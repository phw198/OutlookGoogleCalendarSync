using log4net;
using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace OutlookGoogleCalendarSync {
    internal class ModernToastNotificationHandler : INotificationHandler {
        private static readonly ILog log = LogManager.GetLogger(typeof(ModernToastNotificationHandler));
        private static readonly string instanceId = Guid.NewGuid().ToString("N");
        private static readonly string activationPipeName = "OGCS.Toast." + instanceId;
        private static bool activationCallbacksRegistered = false;
        private static Action<NotificationActivation> activationCallback;
        private static CancellationTokenSource activationPipeCancellation;
        private readonly Action<NotificationActivation> onClick;
        private readonly NotifyIcon ownerIcon;

        public ModernToastNotificationHandler(NotifyIcon ownerIcon, Action<NotificationActivation> onClick) {
            this.ownerIcon = ownerIcon;
            this.onClick = onClick;
            activationCallback = onClick;
            RegisterActivationHandler();
        }

        /// <summary>
        /// Handles application startup activation triggered by a toast notification and processes activation arguments.
        /// </summary>
        /// <remarks>Subscribes to toast activation events, processes activation arguments, and ensures
        /// proper event unsubscription after handling. Waits up to two seconds for activation to complete.</remarks>
        /// <returns>true if the application was activated by a toast notification; otherwise, false.</returns>
        public static bool HandleStartupActivation() {
            try {
                if (!ToastNotificationManagerCompat.WasCurrentProcessToastActivated()) return false;
            } catch (System.Exception ex) {
                ex.LogAsFail().Analyse("Unable to determine whether OGCS was started by a toast activation.");
                return false;
            }

            ManualResetEventSlim activationReceived = new ManualResetEventSlim(false);
            OnActivated activationHandler = args => {
                Dictionary<string, string> arguments = getTriggerArguments(args.Argument);
                if (arguments.TryGetValue("instance", out string targetInstance) && !string.IsNullOrEmpty(targetInstance)) {
                    string action = arguments.TryGetValue("action", out string forwardedAction) ? forwardedAction : string.Empty;
                    string browserUrl = getBrowserUrl(arguments);
                    if (!forwardActivation(targetInstance, new NotificationActivation { Action = action, BrowserUrl = browserUrl }) && !string.IsNullOrEmpty(browserUrl))
                        Helper.OpenBrowser(browserUrl);
                } else {
                    log.Warn("Ignoring startup toast activation without an instance identifier.");
                }
                activationReceived.Set();
            };

            try {
                ToastNotificationManagerCompat.OnActivated += activationHandler;
                activationReceived.Wait(TimeSpan.FromSeconds(2));
            } catch (System.Exception ex) {
                ex.Analyse("Unable to handle toast activation during OGCS startup.");
            } finally {
                ToastNotificationManagerCompat.OnActivated -= activationHandler;
            }

            return true;
        }

        private void RegisterActivationHandler() {
            if (activationCallbacksRegistered) return;

            try {
                startActivationPipe();
                ToastNotificationManagerCompat.OnActivated += args => {
                    Dictionary<string, string> arguments = getTriggerArguments(args.Argument);
                    if (!arguments.TryGetValue("instance", out string targetInstance) || string.IsNullOrEmpty(targetInstance)) {
                        log.Warn("Ignoring toast activation without an instance identifier.");
                        return;
                    }
                    if (targetInstance != instanceId) {
                        forwardActivation(targetInstance, new NotificationActivation {
                            Action = arguments.TryGetValue("action", out string forwardedAction) ? forwardedAction : string.Empty,
                            BrowserUrl = getBrowserUrl(arguments)
                        });
                        return;
                    }
                    string tag = arguments.TryGetValue("action", out string action) ? action : string.Empty;
                    log.Debug("Toast activated: " + tag);
                    activationCallback?.Invoke(new NotificationActivation { Action = tag, BrowserUrl = getBrowserUrl(arguments) });
                };
                activationCallbacksRegistered = true;

            } catch (System.Exception ex) {
                ex.Analyse("Unable to register toast activation callback for Windows notification center clicks.");
            }
        }

        #region Inter-process pipeline communication for toasts clicked from Windows Action Center
        //For multi-instance OGCS, if the wrong process is activated, we need to forward the activation to the correct process.
        //This is done via a named pipe, which is created by the process that created the toast notification.
        private static void startActivationPipe() {
            activationPipeCancellation = new CancellationTokenSource();
            Task.Run(() => listenForActivation(activationPipeCancellation.Token));
        }

        private static async Task listenForActivation(CancellationToken cancellationToken) {
            while (!cancellationToken.IsCancellationRequested) {
                try {
                    using (NamedPipeServerStream server = new NamedPipeServerStream(
                        activationPipeName,
                        PipeDirection.In,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous)) {
                        await server.WaitForConnectionAsync(cancellationToken);
                        using (StreamReader reader = new StreamReader(server)) {
                            string action = await reader.ReadLineAsync();
                            string browserUrl = await reader.ReadLineAsync();
                            if (action != null) activationCallback?.Invoke(new NotificationActivation { Action = action, BrowserUrl = browserUrl });
                        }
                    }
                } catch (OperationCanceledException) {
                    return;
                } catch (System.Exception ex) {
                    ex.Analyse("Unable to receive toast activation for this OGCS instance.");
                }
            }
        }

        private static bool forwardActivation(string targetInstance, NotificationActivation activation) {
            try {
                log.Info("Forwarding toast activation to target OGCS instance: " + targetInstance);
                using (NamedPipeClientStream client = new NamedPipeClientStream(
                    ".",
                    "OGCS.Toast." + targetInstance,
                    PipeDirection.Out,
                    PipeOptions.Asynchronous)) {
                    client.Connect(500);
                    using (StreamWriter writer = new StreamWriter(client)) {
                        writer.AutoFlush = true;
                        writer.WriteLine(activation?.Action ?? string.Empty);
                        writer.WriteLine(activation?.BrowserUrl ?? string.Empty);
                    }
                }
                return true;
            } catch {
                log.Info("Target OGCS instance is no longer running.");
                return false;
            }
        }

        private static Dictionary<string, string> getTriggerArguments(string argument) {
            Dictionary<string, string> arguments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(argument)) return arguments;

            foreach (string part in argument.Split(new[] { '&', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                string[] keyValue = part.Trim().Split(new[] { '=' }, 2);
                if (keyValue.Length == 2 && (new string[] { "action", "instance", "url" }.Contains(keyValue[0]))) {
                    arguments[keyValue[0]] = keyValue[1];
                }
            }
            return arguments;
        }

        private static string getBrowserUrl(Dictionary<string, string> arguments) {
            if (!arguments.TryGetValue("url", out string encodedUrl) || string.IsNullOrEmpty(encodedUrl)) return string.Empty;
            try {
                return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encodedUrl));
            } catch (FormatException) {
                log.Warn("Ignoring malformed browser URL in toast activation.");
                return string.Empty;
            }
        }
        #endregion

        public void ShowNotification(string message, ToolTipIcon iconType = ToolTipIcon.None, string tagValue = "", string browserUrl = "") {
            String title = "Outlook Google Calendar Sync" + (string.IsNullOrEmpty(Program.Title) ? "" : $" - {Program.Title}");
            try {
                String originalMessage = message;
                List<String> messages = message.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                Regex syncSummary = new Regex(@"\d+\screated;\s\d+\sdeleted;\s\d+\supdated");
                if (syncSummary.IsMatch(messages[0])) {
                    if (Sync.Engine.Instance.ActiveProfile is SettingsStore.Calendar profile)
                        title = $"Profile '{profile._ProfileName}' Sync Summary";
                    else
                        title = "Sync Summary";
                } else if (messages.Count > 1) {
                    title = messages[0].TrimEnd('.');
                    messages.RemoveAt(0);
                    message = string.Join("\r\n", messages);
                }

                ToastContent toastContent = new ToastContentBuilder()
                    .AddText(title)
                    .AddText(message)
                    .AddArgument("instance", instanceId)
                    .AddArgument("action", tagValue)
                    .AddArgument("url", string.IsNullOrEmpty(browserUrl) ? string.Empty : Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(browserUrl)))
                    .AddArgument("message", originalMessage)
                    .SetToastScenario(ToastScenario.Default)
                    .GetToastContent();

                XmlDocument toastXml = toastContent.GetXml();
                ToastNotifierCompat notifier = ToastNotificationManagerCompat.CreateToastNotifier();
                ToastNotification toast = new Windows.UI.Notifications.ToastNotification(toastXml);
                toast.Dismissed += (sender, args) => {
                    log.Debug("Toast dismissed");
                };
                notifier.Show(toast);

            } catch (System.Exception ex) {
                ex.Analyse("Unable to show toast");
                ownerIcon?.ShowBalloonTip(500, title, message, iconType);
            }
        }
    }
}
