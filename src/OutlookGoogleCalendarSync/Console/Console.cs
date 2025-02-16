using Ogcs = OutlookGoogleCalendarSync;
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
        private String _content = "";
        private String content {
            get { return _content; }
            set {
                do {
                    _content = value;
                    try {
                        if (this.wb.InvokeRequired) {
                            this.wb.Invoke((MethodInvoker)(() => {
                                wb.DocumentText = _content;
                            }));
                        } else
                            this.wb.DocumentText = _content;
                    } catch (System.Exception ex) {
                        Ogcs.Exception.Analyse(ex);
                    }

                    System.Windows.Forms.Application.DoEvents();
                    while (navigationStatus != NavigationStatus.completed) {
                        System.Threading.Thread.Sleep(250);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    System.Threading.Thread.Sleep(15);
                } while (this.DocumentText != _content);

                System.Windows.Forms.Application.DoEvents();
            }
        }
        public String DocumentText {
            get {
                String documentText = "";
                if (this.wb == null)
                    return null;
                else {
                    if (this.wb.InvokeRequired) {
                        this.wb.Invoke((MethodInvoker)(() => {
                            documentText = wb.DocumentText;
                        }));
                    } else
                        documentText = this.wb.DocumentText;
                }
                return documentText;
            }
        }

        #region Notes
        //If we don't want to depend on the emoji-css project, we could store the images as resources and reference as:
        //  filter: progid:DXImageTransform.Microsoft.AlphaImageLoader(src='file:///C:\Users\Paul\Git\OutlookGoogleCalendarSync\src\OutlookGoogleCalendarSync\bin\Debug\images\warning.png', sizingMethod='scale');

        //Also, the default CSS style "background-size: contain;" is not understood by IE8, which is the version of the webbrowser control
        //The workaround is either to
        // - user meta tag X-UA-Compatible; or
        // - add a registry key in HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION; or
        // - an override css for each emoji using AlphaImageLoader
        #endregion

        #region HTML head code
        private String header = @"
<html>
    <head>
        <meta http-equiv='X-UA-Compatible' content='IE=edge' /> <!-- Make webbrowser control display IE7< content -->
        <!--- <link href='https://afeld.github.io/emoji-css/emoji.css' rel='stylesheet'> -->
        <link href='"+ System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\Console\emoji.css' rel='stylesheet'>
        <style> 
            a.no-decoration {
                text-decoration: none; color: black;
            }
            p {
                margin: auto;
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
            .em-repeat, .em-repeat-one {
                padding-right: 0px;
                margin-left: 0px;
            }
            .config, .info, .error, .warning, .news {
                margin-top: 8px;
                padding-bottom: 4px;
                margin-bottom: 10px;
                padding-left: 5px;
                border-left-width: 10px;
                border-left-style: solid;
                border-bottom-left-radius: 5px;
                border-top-left-radius: 5px;
                margin-left: -15px;
            }
            .config {
                background-color: lightgray;
                border-left-color: yellow;
            }
            .news {
                background-color: #b5f9b5;
                border-left-color: darkseagreen;
            }
            .info {
                background-color: lightblue;
                border-left-color: cornflowerblue;
            }
            .error {
                background-color: pink;
                border-left-color: red;
            }
            .warning {
                background-color: lightgoldenrodyellow;
                border-left-color: yellow;
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
        <script>
            function mp(payload) {
                fetch('https://www.google-analytics.com/mp/collect?api_secret=kWOsAm2tQny1xOjiwMyC5Q&measurement_id=G-S6RMS8GHEE', {
                    mode: 'no-cors',
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                    },
                    body: payload
                    })
               .then(response => response.json())
               .then(response => console.log(JSON.stringify(response)));
            }
        </script>
        <script>
            function toggle() {
                var news = document.getElementById('news');
                var newsText = document.getElementById('newsToggleText');
                if (news.style.display == 'block') {
                    news.style.display = 'none';
                    newsText.innerHTML = ""<a href='#shownews' class='no-decoration'>[+] Show</a>"";
                } else {
                    news.style.display = 'block';
                    newsText.innerHTML = ""<a href='#hidenews' class='no-decoration'>[&#8211] Hide</a>"";
                    scrollToBottom();
                }
            }
        </script>
    </head>
    <body onLoad='scrollToBottom();'>
        <div id='content'>";
        #endregion

        private String footer = @"</div></body></html>";

        private NavigationStatus? navigationStatus = null;
        private enum NavigationStatus {
            navigating,
            navigated,
            completed,
            cleared
        }

        public Console(WebBrowser wb) {
            if (this.wb != null) return;
            this.wb = wb;

            if (Program.InDeveloperMode)
                wb.IsWebBrowserContextMenuEnabled = true;

            wb.Navigate("about:blank");
            wb.AllowNavigation = true;
            wb.WebBrowserShortcutsEnabled = false;
            wb.Document.Write(header + footer);
            this.awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);

            wb.DocumentCompleted += console_DocumentCompleted;
            wb.Navigated += console_Navigated;
            wb.Navigating += console_Navigating;

            awaitRefresh();
            disableClickSounds();
            log.Fine("Console initialised.");
        }

        private void console_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            this.awaitingRefresh = false;
            navigationStatus = NavigationStatus.completed;
            log.UltraFine("Document completed.");
        }

        private void console_Navigating(object sender, WebBrowserNavigatingEventArgs e) {
            if (!Forms.Main.Instance.Visible) return;

            if (e.Url.Scheme != "about") {
                Helper.OpenBrowser(e.Url.OriginalString);
                e.Cancel = true;
                return;
            }
            if (e.Url.Fragment == "#hidenews") {
                Settings.Instance.HideNews = DateTime.UtcNow;
            } else if (e.Url.Fragment == "#shownews") {
                Settings.Instance.HideNews = new DateTime();
            }

            navigationStatus = NavigationStatus.navigating;
            log.UltraFine("Console navigating.");
        }
        private void console_Navigated(object sender, WebBrowserNavigatedEventArgs e) {
            if (this.wb.ReadyState != WebBrowserReadyState.Complete) {
                log.UltraFine("Navigated status = " + this.wb.ReadyState.ToString());
                return;
            }

            navigationStatus = NavigationStatus.navigated;
            log.UltraFine("Console finished navigating");
        }

        private void awaitRefresh() {
            while (this.awaitingRefresh && !isCleared()) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            this.awaitingRefresh = false;
        }

        private Boolean isCleared() {
            return (this.wb.DocumentText == "<HTML></HTML>\0" || this.wb.DocumentText == header + footer);
        }

        public void Clear() {
            if (isCleared()) return;

            content = header + footer;
            awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);
            awaitRefresh();
            navigationStatus = NavigationStatus.cleared;
            log.Fine("Document cleared.");
        }

        public void BuildOutput(string s, ref System.Text.StringBuilder sb, Boolean logit = true) {
            if (logit) log.Debug(s);
            sb.Append("<p>" + s + "</p>");
        }

        public enum Markup {
            calendar,
            checkered_flag,
            config,
            error,
            fail,
            h2,
            info,
            mag_right,
            warning,
            appointmentEnd, //margin top and bottom
            sectionEnd, //Add horizontal rule below the line
            syncDirection
        }

        public void Update(StringBuilder moreOutput, Markup? markupPrefix = null, Boolean verbose = false, bool notifyBubble = false, Boolean logit = false) {
            Update(moreOutput.ToString(), markupPrefix, newLine: false, verbose: verbose, notifyBubble: notifyBubble, logit: logit);
        }

        /// <summary>
        /// Log a different string than that displayed to the Console
        /// </summary>
        /// <param name="moreOutput">Console output</param>
        /// <param name="logEntry">Log output</param>
        public void Update(String moreOutput, String logEntry, Markup? markupPrefix = null, bool newLine = true, Boolean verbose = false, bool notifyBubble = false) {
            if (string.IsNullOrEmpty(logEntry))
                Update(moreOutput, markupPrefix, newLine, verbose, notifyBubble);
            else {
                Update(moreOutput, markupPrefix, newLine, verbose, notifyBubble, logit: false);
                logLinesSansHtml(logEntry, markupPrefix, verbose);
            }
        }

        /// <summary>
        /// Update the console with further text
        /// </summary>
        /// <param name="moreOutput">The text to update the console with</param>
        /// <param name="markupPrefix">The type of update, may cause emoji prefix</param>
        /// <param name="newLine">Append a new line at the end</param>
        /// <param name="verbose">Only output if verbose is set 'On'</param>
        /// <param name="notifyBubble">Trigger a system bubble for notification</param>
        /// <param name="logit">Send the text to logfile</param>
        public void Update(String moreOutput, Markup? markupPrefix = null, bool newLine = true, Boolean verbose = false, bool notifyBubble = false, Boolean logit = true) {
            ///Accessing the DOM is terribly slow, so not using this method.
            ///HtmlDocument doc = Forms.Main.Instance.GetControlPropertyThreadSafe(this.wb, "Document") as HtmlDocument;
            ///HtmlElement element = doc.GetElementById("content");
            ///HtmlElement element = doc.All["content"]; //Slightly faster

            if (Forms.Main.Instance.IsDisposed) return;

            if ((verbose && Settings.Instance.VerboseOutput) || !verbose) {
                //Let's grab the 'content' div with regex
                Regex rgx = new Regex("<div id=\'content\'>(.*)</div>", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                MatchCollection matches = rgx.Matches(content);

                String contentInnerHtml = "";
                if (matches != null && matches.Count > 0) {
                    contentInnerHtml = matches[0].Result("$1");
                }

                moreOutput = moreOutput.Replace("\r\n", "<br/>");
                String htmlOutput = parseEmoji(moreOutput, markupPrefix);

                if (logit) logLinesSansHtml(htmlOutput, markupPrefix, verbose);

                //Don't add append line break to Markup that's already wrapped in <div> tags
                if (markupPrefix != null && (new Markup[] { Markup.info, Markup.warning, Markup.fail, Markup.error }.ToList()).Contains((Markup)markupPrefix))
                    newLine = false;

                contentInnerHtml += htmlOutput + (newLine ? "<br/>" : "");
                content = header + contentInnerHtml + footer;

                if (Forms.Main.Instance.NotificationTray != null && notifyBubble) {
                    Forms.Main.Instance.NotificationTray.ShowBubbleInfo("Issue encountered.\n" +
                        "Please review output on the main 'Sync' tab", ToolTipIcon.Warning);
                }
            }
        }

        public void UpdateWithError(String moreOutput, System.Exception ex, bool notifyBubble = false, String logEntry = null) {
            Markup emoji = Markup.error;
            if (ex.LoggingAsFail())
                emoji = Markup.fail;
            Update(moreOutput + (!string.IsNullOrEmpty(moreOutput) ? "<br/>" : "") + Ogcs.Exception.FriendlyMessage(ex), logEntry, emoji, notifyBubble: notifyBubble);
        }

        /// <summary>Log the output sans HTML tags.</summary>
        private void logLinesSansHtml(String htmlOutput, Markup? markupPrefix = null, Boolean verbose = false) {
            String tagsStripped = Regex.Replace(htmlOutput, "(</p>|<br/?>)", "\r\n");
            tagsStripped = Regex.Replace(tagsStripped, "<span class='em em-repeat'></span>", "(R)");
            tagsStripped = Regex.Replace(tagsStripped, "<span class='em em-repeat-one'></span>", "(R1)");
            tagsStripped = Regex.Replace(tagsStripped, "<.*?>", String.Empty);
            String[] logLines = tagsStripped.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (markupPrefix == Markup.warning)
                logLines.ToList().ForEach(l => log.Warn(l));
            else if (markupPrefix == Markup.fail)
                logLines.ToList().ForEach(l => log.Fail(l));
            else if (markupPrefix == Markup.error)
                logLines.ToList().ForEach(l => log.Error(l));
            else if (verbose)
                logLines.ToList().ForEach(l => log.Debug(l));
            else
                logLines.ToList().ForEach(l => log.Info(l));
        }

        private String parseEmoji(String output, Markup? markupPrefix = null) {
            if (markupPrefix != null) output = (":" + markupPrefix.ToString() + ":") + output;

            try {
                //div
                output = Regex.Replace(output, ":info:(<p>)*", "<div class='info'>$1<span class='em em-information_source'></span>");
                output = Regex.Replace(output, ":warning:(<p>)*", "<div class='warning'>$1<span class='em em-warning'></span>");
                output = Regex.Replace(output, ":(error|fail):(<p>)*", "<div class='error'>$2<span class='em em-collision'></span>");
                output = Regex.Replace(output, ":config:(<p>)*", "<div class='config'>$1<span class='em em-gear'></span>");
                output = Regex.Replace(output, ":newspaper:(<p>)*", "<div class='news'>$1<span class='em'></span>");
                if (output.StartsWith("<div")) output += "</div>";

                Regex rgx = new Regex(":clock(\\d{1,4}):<p>", RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(output);
                if (matches.Count > 0) {
                    String clockTime = matches[0].Result("$1");
                    output = output.Replace(":clock" + clockTime + ":<p>", "<div class='info'><p><span class='em em-clock" + clockTime + "'></span>") + "</div>";
                }

                //h2
                output = output.Replace(":h2:", "<h2 class='sectionHeader'>");
                output = output.Replace(":mag_right:", "<h2 class='sectionHeader'><span class='em em-mag_right'></span>");
                output = output.Replace(":checkered_flag:", "<h2 class='sectionHeader'><span class='em em-checkered_flag'></span>");
                output = output.Replace(":syncDirection:", "<h2 class='sectionHeader'><span class='em em-repeat' style='padding-right: 9px; margin-left:-30px;'></span>");
                if (output.StartsWith("<h2")) output += "</h2>";

                //sectionEnd
                output = output.Replace(":sectionEnd:", "<p class='sectionEnd'>");
                if (output.StartsWith("<p")) output += "</p>";

                output = output.Replace(":appointmentEnd:", "<p class='appointmentEnd'>");
                if (output.StartsWith("<p")) output += "</p>";

                output = output.Replace(":calendar:", "<span class='em em-date' style='margin-top:5px'></span>");
                output = output.Replace("(R)", "<span class='em em-repeat'></span>");
                output = output.Replace("(R1)", "<span class='em em-repeat-one'></span>");
                output = output.Replace("=> ", "");

            } catch (System.Exception ex) {
                log.Error("Failed parsing for emoji.");
                Ogcs.Exception.Analyse(ex);
            }
            return output;
        }

        public void FormatEventChanges(StringBuilder sb, String anonymised) {
            sb.Insert(0, ":" + Markup.calendar.ToString() + ":");

            String[] lines = sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (lines.Count() == 1) return;

            StringBuilder table = new StringBuilder("<table class='eventChanges' border='0'>");
            table.Append("<tr><th class='eventChanges'>Attribute</th><th class='eventChanges'>Change</th></tr>");
            for (int l = 1; l < lines.Count(); l++) {
                String newRow = "<tr>";
                newRow += Regex.Replace(lines[l], @"^(\w+|\w+[\s/-]\w+|(Attendee|Organiser) (added|updated|removed|.*?Status|.*?Optional Check)):\s*", "<td class='eventChanges'>$1</td><td>");
                newRow = newRow.Replace("=>", "→");
                table.Append(newRow + "</td></tr>");
            }
            table.Append("</table>");

            if (Settings.Instance.AnonymiseLogs) {
                MatchCollection matches = Regex.Matches(anonymised, @"^Subject:\s(.*?)\s=>\s(.*?)$", RegexOptions.Multiline);
                if (matches.Count > 0) {
                    anonymised = anonymised.Replace(matches[0].Value, "Subject: " + Ogcs.Google.Authenticator.GetMd5(matches[0].Groups[1].Value) + " => " + Ogcs.Google.Authenticator.GetMd5(matches[0].Groups[2].Value.TrimEnd("\r".ToCharArray())));
                }
            }
            Update(lines[0] + "<br/>" + table.ToString(), anonymised, verbose: true, newLine: false);
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
                    Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath + ".Current", true).SetValue(defaultKeyName, "");
                } else {
                    log.Fine("Unmuting navigation click sounds.");
                    soundRegKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath + ".Default", false);
                    if (soundRegKey != null) {
                        String defaultSound = soundRegKey.GetValue(defaultKeyName, "", RegistryValueOptions.DoNotExpandEnvironmentNames) as String;
                        if (string.IsNullOrEmpty(defaultSound))
                            log.Warn("No default navigation sound found to reinstate as current.");
                        else
                            Microsoft.Win32.Registry.CurrentUser.OpenSubKey(navigatingKeyPath + ".Current", true).SetValue(defaultKeyName, defaultSound);
                    } else
                        log.Warn("Could not find default navigation sound registry key.");
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }
        #endregion

        /// <summary>Invoke execution of a Google App Script via a GET request</summary>
        public void CallGappScript() {
            /*
            try {
                Forms.Main.Instance.GappBrowser.Navigate("https://script.google.com/macros/s/<GUID>/exec?key=value");
                while (Forms.Main.Instance.GappBrowser.ReadyState != WebBrowserReadyState.Complete) {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
            log.Debug("Done");
            */
        }
    }
}
