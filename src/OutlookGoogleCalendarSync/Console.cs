using log4net;
using Microsoft.Win32;
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    public class Console {
        private static readonly ILog log = LogManager.GetLogger(typeof(Console));

        private WebBrowser wb;
        private Boolean awaitingRefresh;
        
        #region Notes
        //If we don't want to depend on the emoji-css project, we /could/ store the images as resources and reference as:
        //  filter: progid:DXImageTransform.Microsoft.AlphaImageLoader(src='file:///C:\Users\Paul\Git\OutlookGoogleCalendarSync\src\OutlookGoogleCalendarSync\bin\Debug\images\warning.png', sizingMethod='scale');

        //Also, the default CSS style "background-size: contain;" is not understood by IE8, which is the version of the webbrowser control
        //The workaround is either to
        // - user meta tag X-UA-Compatible; or
        // - add a registry key in HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION; or
        // - an override css for each emoji using AlphaImageLoader
        #endregion

        private String header = @"
<html>
    <head>
        <meta http-equiv='X-UA-Compatible' content='IE=edge' /> <!-- Make webbrowser control display IE7< content -->
        <link href='https://afeld.github.io/emoji-css/emoji.css' rel='stylesheet'>
        <style> 
            p {
                margin: auto;
                padding-left: 5px;
            }
            div#content {
                font-family: Arial;
                font-size: 14px;
                margin-left: 27px;
            }
            .em {
                vertical-align: text-bottom;
                background-position: left;
                padding-right: 9px;
                margin-left: -30px;
            }
         	.em-repeat {
                padding-right: 0px;
				margin-left: 0px;
			}
            .info, .error {
                margin-top: 8px;
				padding-bottom: 4px;
				margin-bottom: 10px;
				border-left-width: 10px;
				border-left-style: solid;
				border-bottom-left-radius: 5px;
				border-top-left-radius: 5px;
				margin-left: -15px;
			}
            .info {
				background-color: lightblue;
				border-left-color: cornflowerblue;
			}
            .error {
				background-color: pink;
				border-left-color: red;
			}
			h2.sectionHeader {
                font-size: 14px;
                margin-bottom: 3px;
            }
			.sectionEnd {
				padding-left: 0px;
				border-bottom: 1px solid lightgray;
				padding-bottom: 6px;
				margin-bottom: 10px;
			}
            .appointmentStart {
                margin-top: 5px;
            }
            .appointmentEnd {
                margin-bottom: 10px;
                margin-top: 5px;
                padding-left: 0px;
            }
			table.eventChanges {
                font-size: 14px;
                margin-top: 5px;
                border-spacing: 4px;
                margin-left: 10px;
			}
            th.eventChanges {
                text-align: left;
                font-weight: normal;
                background-color: #d6edff;
                margin: 10px;
                border-bottom: gray 1px dashed;
            }
			td.eventChanges {
				padding-right: 10px;
				vertical-align: top;
			}
            tr:nth-child(odd) {
                background-color: #eae9e9;
            }
            <!-- This is the other way of overriding IE7 not supporting background-size
            .em-warning {
                filter: progid:DXImageTransform.Microsoft.AlphaImageLoader(src='https://afeld.github.io/emoji-css/emoji/warning.png', sizingMethod='scale');
                background-image: none;
            }
            -->
        </style>
        <script>
          function scrollToBottom() {
            window.scrollTo(0,document.body.scrollHeight);
          }
        </script>
    </head>
    <body onLoad='scrollToBottom();'>
        <div id='content'>";
        
        private String footer = @"</div></body></html>";
        private Boolean isNavigating = false;
        private DateTime startedNavigating;

        public Console(WebBrowser wb) {
            if (this.wb != null) return;
            this.wb = wb;

            wb.Navigate("about:blank");
            wb.AllowNavigation = true;
            wb.Document.Write(header + footer);
            this.awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);

            wb.DocumentCompleted += console_DocumentCompleted;
            wb.Navigated += console_Navigated;
            wb.Navigating += console_Navigating;
            
            awaitRefresh();
            disableClickSounds();
            log.Debug("Console initialised.");
        }

        private void console_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            wb.DocumentCompleted -= console_DocumentCompleted;
            this.awaitingRefresh = false;
            isNavigating = false;
        }

        private void console_Navigated(object sender, WebBrowserNavigatedEventArgs e) {
            if (this.wb.ReadyState != WebBrowserReadyState.Complete) {
                log.Debug("Navigated, but not completely ready yet.");
                return;
            }
            log.Debug("Console finished navigating");
            isNavigating = false;
        }
        private void console_Navigating(object sender, WebBrowserNavigatingEventArgs e) {
            log.Debug("Console started navigating");
            isNavigating = true;
            startedNavigating = DateTime.Now;
        }

        private void awaitRefresh() {
            while (this.awaitingRefresh && !isCleared()) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            this.awaitingRefresh = false;
            isNavigating = false;
        }

        private Boolean isCleared() {
            return (this.wb.DocumentText == "<HTML></HTML>\0" || this.wb.DocumentText == header + footer);
        }

        public void Clear() {
            if (isCleared()) return;
            
            this.wb.DocumentText = header + footer;
            awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);
            awaitRefresh();
        }

        public void BuildOutput(string s, ref System.Text.StringBuilder sb, Boolean logit = true) {
            if (logit) log.Debug(s);
            sb.Append("<p>" + s + "</p>");
        }

        public enum Markup {
            checkered_flag,
            error,
            h2,
            info,
            mag_right,
            sectionEnd, //Add horizontal rule below the line
            warning
        }

        public void Update(StringBuilder moreOutput, Markup? markupPrefix = null, Boolean verbose = false, bool notifyBubble = false, Boolean logit = false) {
            Update(moreOutput.ToString(), markupPrefix, newLine: false, verbose: verbose, notifyBubble: notifyBubble, logit: logit);
        }

        public void Update(String moreOutput, Markup? markupPrefix = null, bool newLine = true, Boolean verbose = false, bool notifyBubble = false, Boolean logit = true) {
            ///Accessing the DOM is terribly slow, so not using this method.
            ///HtmlDocument doc = MainForm.Instance.GetControlPropertyThreadSafe(this.wb, "Document") as HtmlDocument;
            ///HtmlElement element = doc.GetElementById("content");
            ///HtmlElement element = doc.All["content"]; //Slightly faster

            if ((verbose && Settings.Instance.VerboseOutput) || !verbose) {
                while (isNavigating && startedNavigating > DateTime.Now.AddSeconds(-2)) {
                    System.Threading.Thread.Sleep(250);
                    System.Windows.Forms.Application.DoEvents();
                }

                //Let's grab the 'content' div with regex
                String allDocument = MainForm.Instance.GetControlPropertyThreadSafe(this.wb, "DocumentText") as String;
                Regex rgx = new Regex("<div id=\'content\'>(.*)</div>", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                MatchCollection matches = rgx.Matches(allDocument);

                String contentInnerHtml = "";
                if (matches.Count == 0) {
                    log.Error("empty doc!");
                    System.Windows.Forms.Application.DoEvents();
                }
                if (matches.Count > 0) {
                    contentInnerHtml = matches[0].Result("$1");
                }

                String htmlOutput = parseEmoji(moreOutput, markupPrefix);

                if (logit) {
                    //Log the output sans HTML tags
                    String tagsStripped = Regex.Replace(htmlOutput, "(</p>|<br/?>)", "\r\n");
                    tagsStripped = Regex.Replace(tagsStripped, "<.*?>", String.Empty);
                    String[] logLines = tagsStripped.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    if (markupPrefix == Markup.warning)
                        logLines.ToList().ForEach(l => log.Warn(l));
                    else if (markupPrefix == Markup.error)
                        logLines.ToList().ForEach(l => log.Error(l));
                    else if (verbose)
                        logLines.ToList().ForEach(l => log.Debug(l));
                    else
                        logLines.ToList().ForEach(l => log.Info(l));
                }

                contentInnerHtml += htmlOutput + (newLine ? "<br/>" : "");
                this.wb.DocumentText = header + contentInnerHtml + footer;
                System.Windows.Forms.Application.DoEvents();
                
                if (MainForm.Instance.NotificationTray != null && notifyBubble & Settings.Instance.ShowBubbleTooltipWhenSyncing) {
                    MainForm.Instance.NotificationTray.ShowBubbleInfo("Issue encountered.\n" +
                        "Please review output on the main 'Sync' tab", ToolTipIcon.Warning);
                }
            }
        }

        private String parseEmoji(String output, Markup? markupPrefix = null) {
            if (markupPrefix != null) output = (":" + markupPrefix.ToString() + ":") + output;

            try {
                //div
                output = output.Replace(":info:<p>", "<div class='info'><p><span class='em em-information_source'></span>");
                output = output.Replace(":warning:<p>", "<div class='warning'><p><span class='em em-warning'></span>");
                output = output.Replace(":error:<p>", "<div class='error'><p><span class='em em-collision'></span>");
                if (output.StartsWith("<div")) output += "</div>";

                Regex rgx = new Regex(":clock(\\d{1,4}):<p>", RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(output);
                if (matches.Count > 0) {
                    String clockTime = matches[0].Result("$1");
                    output = output.Replace(":clock"+ clockTime +":<p>", "<div class='info'><p><span class='em em-clock" + clockTime + "'></span>") + "</div>";
                }

                //h2
                output = output.Replace(":h2:", "<h2 class='sectionHeader'>");
                output = output.Replace(":mag_right:", "<h2 class='sectionHeader'><span class='em em-mag_right'></span>");
                output = output.Replace(":checkered_flag:", "<h2 class='sectionHeader'><span class='em em-checkered_flag'></span>");
                if (output.StartsWith("<h2")) output += "</h2>";

                //sectionEnd
                output = output.Replace(":sectionEnd:", "<p class='sectionEnd'>");
                if (output.StartsWith("<p")) output += "</p>";

                output = output.Replace(":appointmentEnd:", "<p class='appointmentEnd'>");
                if (output.StartsWith("<p")) output += "</p>";
                
                output = output.Replace(":date:", "<span class='em em-date' style='margin-top:5px'></span>");
                output = output.Replace("(R)", "<span class='em em-repeat'></span>");
                output = output.Replace("=>", "");

            } catch (System.Exception ex) {
                log.Error("Failed parsing for emoji.");
                OGCSexception.Analyse(ex);
            }
            return output;
        }

        public void FormatEventChanges(StringBuilder sb) {
            sb.Insert(0, ":date:");

            String[] lines = sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (lines.Count() == 1) return;

            StringBuilder table = new StringBuilder("<table class='eventChanges' border='0'>");
            table.Append("<tr><th class='eventChanges'>Attribute</th><th class='eventChanges'>Change</th></tr>");
            for (int l = 1; l < lines.Count(); l++) {
                String newRow = "<tr>";
                newRow += Regex.Replace(lines[l], "^(\\w+|(Start|End) \\w+|Attendee.*?Status|Reminder Default):\\s*", "<td class='eventChanges'>$1</td><td>");
                newRow = newRow.Replace("=>", "→");
                table.Append(newRow + "</td></tr>");
            }
            table.Append("</table>");

            Update(lines[0] +"<br/>"+ table.ToString(), verbose: true, newLine: false);
        }

        #region Mute webbrowser navigation click sounds
        const int FEATURE_DISABLE_NAVIGATION_SOUNDS = 21;
        const int SET_FEATURE_ON_THREAD = 0x00000001;
        const int SET_FEATURE_ON_PROCESS = 0x00000002;
        const int SET_FEATURE_IN_REGISTRY = 0x00000004;
        const int SET_FEATURE_ON_THREAD_LOCALMACHINE = 0x00000008;
        const int SET_FEATURE_ON_THREAD_INTRANET = 0x00000010;
        const int SET_FEATURE_ON_THREAD_TRUSTED = 0x00000020;
        const int SET_FEATURE_ON_THREAD_INTERNET = 0x00000040;
        const int SET_FEATURE_ON_THREAD_RESTRICTED = 0x00000080;

        [System.Runtime.InteropServices.DllImport("urlmon.dll")]
        [System.Runtime.InteropServices.PreserveSig]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Error)]
        static extern int CoInternetSetFeatureEnabled(
            int FeatureEntry,
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] int dwFlags,
            bool fEnable);

        private static void disableClickSounds() {
            int result = CoInternetSetFeatureEnabled(
                FEATURE_DISABLE_NAVIGATION_SOUNDS,
                SET_FEATURE_ON_PROCESS,
                true);
        }

        /// <summary>
        /// The urlmon DLL method isn't 100% reliable.
        /// The fall-back method is to temporarily remove the click registry sound during a sync, 
        /// though this will affect IE navigation outside of OGCS too.
        /// </summary>
        public static void MuteClicks(Boolean mute) {
            RegistryKey soundRegKey = null;
            const String navigatingKeyPath = @"AppEvents\Schemes\Apps\Explorer\Navigating\";
            const String defaultKeyName = "";
            try {
                if (mute) {
                    log.Fine("Muting navigation click sounds.");
                    Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath +".Current", true).SetValue(defaultKeyName, "");
                } else {
                    log.Fine("Unmuting navigation click sounds.");
                    soundRegKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath +".Default", false);
                    if (soundRegKey != null) {
                        String defaultSound = soundRegKey.GetValue(defaultKeyName) as String;
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath +".Current", true).SetValue(defaultKeyName, defaultSound);
                    } else
                        log.Warn("Could not find default navigation sound.");
                }
            } catch (System.Exception ex) {
                log.Error(ex.Message);
            }
        }
        #endregion
    }
}
