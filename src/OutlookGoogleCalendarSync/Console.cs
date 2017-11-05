using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;
using System.Text.RegularExpressions;

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

        String header = @"
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
        
        String footer = @"</div></body></html>";

        public Console(WebBrowser wb) {
            if (this.wb != null) return;
            this.wb = wb;

            wb.Navigate("about:blank");
            wb.AllowNavigation = true;
            wb.Document.Write(header + footer);
            this.awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);

            wb.DocumentCompleted += console_DocumentCompleted;
            
            awaitRefresh();
            log.Debug("Console initialised.");
        }

        private void console_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            this.awaitingRefresh = false;
            wb.DocumentCompleted -= console_DocumentCompleted;
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
            
            this.wb.DocumentText = header + footer;
            awaitingRefresh = true;
            wb.Refresh(WebBrowserRefreshOption.Completely);
            awaitRefresh();
        }

        public void BuildOutput(string s, ref System.Text.StringBuilder sb) {
            log.Debug(s);
            sb.Append("<p>" + s + "</p>");
        }

        public void Update(StringBuilder moreOutput, Boolean verbose = false) {
            Update(moreOutput.ToString(), newLine: false, verbose: verbose);
        }

        public void Update(String moreOutput, bool newLine = true, Boolean verbose = false, bool notifyBubble = false, Boolean logit = false) {
            ///Accessing the DOM is terribly slow, so not using this method.
            ///HtmlDocument doc = MainForm.Instance.GetControlPropertyThreadSafe(this.wb, "Document") as HtmlDocument;
            ///HtmlElement element = doc.GetElementById("content");
            ///HtmlElement element = doc.All["content"]; //Slightly faster

            if ((verbose && Settings.Instance.VerboseOutput) || !verbose) {
                //Let's grab the 'content' div with regex
                String allDocument = MainForm.Instance.GetControlPropertyThreadSafe(this.wb, "DocumentText") as String;
                Regex rgx = new Regex("<div id=\'content\'>(.*)</div>", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                MatchCollection matches = rgx.Matches(allDocument);

                String contentInnerHtml = "";
                if (matches.Count > 0) {
                    contentInnerHtml = matches[0].Result("$1");
                }

                //No new line if an emoji div
                Match match = Regex.Match(moreOutput, "^:\\w+:");
                newLine = !match.Success;

                String htmlOutput = parseEmoji(moreOutput);

                if (logit) {
                    //Log the output sans HTML tags
                    String tagsStripped = Regex.Replace(htmlOutput, "<.*?>", String.Empty);
                    if (htmlOutput.Contains("em-warning"))
                        log.Warn(tagsStripped);
                    else if (htmlOutput.Contains("em-collision"))
                        log.Error(tagsStripped);
                    else if (verbose)
                        log.Debug(tagsStripped);
                    else
                        log.Info(tagsStripped);
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

        private String parseEmoji(String output) {
            try {
                //div
                output = output.Replace(":info:<p>", "<div class='info'><p><span class='em em-information_source'></span>");
                output = output.Replace(":warning:", "<div class='warning'><p><span class='em em-warning'></span>");
                output = output.Replace(":error:", "<div class='error'><p><span class='em em-collision'></span>");
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

                output = output.Replace(":date:", "<span class='em em-date'></span>");
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
            for (int l=1; l<lines.Count(); l++) {
                String newRow = "<tr>";
                newRow = Regex.Replace(lines[l], "^(\\w+):\\s*", "<td>$1</td><td>");
                newRow = newRow.Replace("=>", "→");
                table.Append(newRow + "</td></tr>");
            }
            table.Append("</table>");

            Update(table.ToString(), verbose: true);
        }
    }
}
