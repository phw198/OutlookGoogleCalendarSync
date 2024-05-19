using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.Outlook {
    public class GMeetLogo {
        private static GMeetLogo instance;
        private static readonly ILog log = LogManager.GetLogger(typeof(GMeetLogo));

        public static GMeetLogo Instance {
            get { return instance ??= new GMeetLogo(); }
        }
        public GMeetLogo() {
            GMeetLogoBase64 = base64encode(Properties.Resources.gmeet_logo);
        }

        public String GMeetLogoBase64 {
            get; internal set;
        }

        private String base64encode(System.Drawing.Image img) {
            try {
                byte[] imgBytes = null;
                using (var stream = new System.IO.MemoryStream()) {
                    img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    imgBytes = stream.ToArray();
                }
                return Convert.ToBase64String(imgBytes);

            } catch (System.Exception ex) {
                ex.Analyse("Could not load GMeet logo.");
                return "";
            }
        }
    }

    public static class GMeet {
        private static readonly ILog log = LogManager.GetLogger(typeof(GMeet));

        private const String meetingIdToken = "GMEETURL";
        private const String meetingLogoToken = "GMEETLOGO";
        private static String plainInfo = "\r\nGoogle Meet joining information\r\nGMEETURL\r\nFirst time using Meet?  Learn more  <https://gsuite.google.com/learning-center/products/meet/get-started/>  \r\n\r\n";
        private static String plainHtmlInfo = "Google Meet joining informationGMEETURL <GMEETURL> First time using Meet? Learn more <https://gsuite.google.com/learning-center/products/meet/get-started/>    ";

        /// <summary>
        /// RTF document code for Google Meet details
        /// </summary>
        private static readonly String rtfHeader = @"{\rtf1\ansi\ansicpg1252\deff0\deflang2057{\fonttbl{\f0\fnil\fcharset0 Calibri;}{\f1\fswiss\fprq2\fcharset0 Calibri;}}";
        private static readonly String rtfInfo =
        #region RTF document
            @"
{\colortbl ;\red0\green0\blue255;\red5\green99\blue193;}
{\*\generator Msftedit 5.41.21.2510;}\viewkind4\uc1\pard\lang9\f0\fs22{\pict\wmetafile8\picw2117\pich794\picwgoal1200\pichgoal450 
010009000003480e00000000320e00000000050000000b0200000000050000000c021a03450832
0e0000430f2000cc0000001e005000000000001a0345080000000028000000500000001e000000
0100180000000000201c0000c40e0000c40e00000000000000000000ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
fffffffffffffffffffffffffffffffffffffffcf5efe79f60e38c40e38c40e38c40e38c4075c0
4075c04075c04075c04075c04075c04075c04075c04075c04075c04075c04080c650dcefcfffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffe79f60da6600da6600da6600da66
00da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe38c40da6600da66
00da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac00ffffffffffffffffffaeda8fc5e4afffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe38c
40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac00fffffff3f9ef69bb3047ac0075c040ffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe2e1e0d9d7d6f5f5
f4ffffffffffffffffffffffffffffffffffffffffffffffffecebead9d7d6f5f5f4ffffffffff
ffffffffecebeabcbab8b3b0aeb3b0aecfceccffffffffffffffffffffffffffffffecebeabcba
b8b3b0aeb3b0aecfceccffffffffffffffffffffffffffffffecebeabcbab8bcbab8e2e1e0ffff
ffffffffe38c40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac0047ac0047ac00dcefcf52b11047ac0047ac0075c040ffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a
8768635fd9d7d6ffffffffffffffffffffffffffffffffffffffffffffffffb3b0ae68635fd9d7
d6ffffffffffffc6c4c268635f68635f68635f68635f68635f8d8a87f5f5f4ffffffffffffc6c4
c2716c6968635f68635f68635f68635f84807df5f5f4ffffffffffffe2e1e0716c6968635f6863
5faaa7a5ffffffffffffe38c40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac003f9f0047ac0047ac0047ac0047ac
0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffffffffffffffffffffffffb3b0
ae68635fd9d7d6ffffffd9d7d668635f68635fa09d9bd9d7d6c6c4c27a767368635f8d8a87ffff
ffe2e1e068635f68635f979391d9d7d6c6c4c284807d68635f84807dffffffffffff9793916863
5f84807db3b0aeecebeafffffffffffff69d55f47d1df47d1df47d1df47d1df47d1dd0d396d0ea
bfd0eabfd0eabfd0eabfd0eabfd0eabf69bb3047ac0047ac0047ac003792002d830047ac0047ac
0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffbcbab8cfceccffffffffff
ffffffffb3b0ae68635fd9d7d6ffffff97939168635fb3b0aefffffffffffffffffff5f5f48480
7dc6c4c2ffffff97939168635faaa7a5fffffffffffffffffff5f5f48d8a87c6c4c2ffffffffff
ff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc8426fc8426fc8426fc8426fc84
26fee0c8ffffffffffffffffffffffffffffffffffff75c04047ac0045a900328b002d83002d83
0047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffcfcecc68635f716c
69f5f5f4ffffffffffffb3b0ae68635fd9d7d6ffffff68635f68635ff5f5f4ffffffffffffffff
ffffffffffffffffffffffffff716c6968635fecebeaffffffffffffffffffffffffffffffffff
ffffffffffffff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc8426fc8426fc84
26fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffffff75c04041a2002f86002d83
002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffff7a76
7368635f68635faaa7a5ffffffffffffb3b0ae68635fd9d7d6ffffff68635f68635f68635f6863
5f68635f68635f68635f68635f68635fd9d7d668635f68635f68635f68635f68635f68635f6863
5f68635f68635fd9d7d6ffffff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc84
26fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffffff6bb1402d83
002d83002d83002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffff
ffb3b0ae68635f84807d716c6968635fecebeaffffffb3b0ae68635fd9d7d6ffffff68635f6863
5f84807d8d8a878d8a878d8a878d8a8768635f68635fffffff7a767368635f84807d8d8a878d8a
878d8a878d8a8768635f68635fe2e1e0ffffff8d8a8768635fd9d7d6ffffffffffffffffffffff
fffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffff
ff51b79f2d83002d83002d83002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a876863
5fd9d7d6f5f5f4716c6968635fe2e1e0b3b0ae68635f8d8a87ffffffb3b0ae68635fd9d7d6ffff
ff97939168635fa09d9bffffffffffffffffffc6c4c268635f7a7673ffffffa09d9b68635f9793
91ffffffffffffffffffd9d7d668635f716c69ffffffffffff8d8a8768635fd9d7d6ffffffffff
fffffffffffffffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffff
ffffffffffffff40cbff08b0cf2a86102d83002d83002d830047ac0047ac0047ac0047ac0075c0
40ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ff8d8a8768635fd9d7d6a09d9b68635fa09d9bfffffff5f5f47a767368635fd9d7d6b3b0ae6863
5fd9d7d6ffffffe2e1e068635f68635f8d8a87b3b0aea09d9b68635f68635fcfceccffffffeceb
ea716c6968635f8d8a87b3b0aeaaa7a5716c6968635fbcbab8e2e1e08d8a87716c6968635f8480
7d8d8a87c6c4c2fffffffffffffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffff
ffffffffffffffffffffffffff40cbff00baff08b0cf258d302d83002d830047ac0047ac0047ac
0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffff8d8a8768635fbcbab868635f716c69f5f5f4ffffffffffffc6c4c268635f7a76
73b3b0ae68635fd9d7d6ffffffffffffcfcecc716c6968635f68635f68635f68635fb3b0aeffff
ffffffffffffffe2e1e07a767368635f68635f68635f68635faaa7a5ffffffd9d7d668635f6863
5f68635f68635f68635fb3b0aeffffffffffffd79681ca7457ca7457ca7457ca7457ca74577fcd
e47fdcff7fdcff7fdcff7fdcff7fdcff7fdcff20c2ff00baff00baff03b7ef2291402d830047ac
0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffff8d8a8768635f716c6968635fb3b0aeffffffffffffffffffffff
ff8d8a8768635f7a767368635fd9d7d6fffffffffffffffffff5f5f4d9d7d6b3b0aecfcecceceb
eaffffffffffffffffffffffffffffffffffffd9d7d6b3b0aecfceccecebeaffffffffffffffff
ffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffd8dbfb414eeb3543ea3543ea3543
ea3543ea00baff00baff00baff00baff00baff00baff00baff00baff00baff00baff00baff00ba
ff199b7047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffff8d8a8768635f68635f7a7673ffffffffffffffff
ffffffffffffffd9d7d668635f68635f68635fd9d7d6ffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffffffd8dbfb414e
eb3543ea3543ea3543ea00baff00baff00baff00baff00baff00baff00baff00baff00baff00ba
ff00baff00baff00baffaeda8f47ac0047ac0047ac0075c040ffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635f68635fcfceccffff
ffffffffffffffffffffffffffffffffa09d9b68635f68635fd9d7d6ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffff
ffffffffd8dbfb414eeb3543ea3543ea00baff00baff00baff00baff00baff00baff00baff00ba
ff00baff00baff00baff00baff00baffffffffdcefcf52b11047ac0075c040ffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635f8d8a
87ffffffffffffffffffffffffffffffffffffffffffecebea716c6968635fd9d7d6ffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffd8dbfb414eeb3543ea00baff00baff00baff00baff00baff00ba
ff00baff00baff00baff00baff00baff00baff00bafffffffffffffff3f9ef80c650a2d57fffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe2e1
e0d9d7d6ecebeaffffffffffffffffffffffffffffffffffffffffffffffffe2e1e0d9d7d6f5f5
f4ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffd8dbfb414eeb00baff00baff00baff00ba
ff00baff00baff00baff00baff00baff00baff00baff00baff40cbffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffd8dbfb40cbff40cb
ff40cbff40cbff40cbff40cbff40cbff40cbff40cbff40cbff40cbff60d3ffeffaffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffff030000000000
}\lang2057\f1\par
\b Google Meet joining information\b0\par
{\field{\*\fldinst{HYPERLINK ""GMEETURL""}}{\fldrslt{\ul\cf1 GMEETURL}}}\f1\fs22\par
First time using Meet?{\field{\*\fldinst{ HYPERLINK ""https://gsuite.google.com/learning-center/products/meet/get-started/"" \\\\l ""!/section-2-3?hl=en-GB"" \\\\t ""_blank"" } } {\fldrslt{\cf2\ul\b\~Learn more} } }\cf0\ulnone\b0\f1\fs22\par
\par
\pard\sa200\sl276\slmult1\lang9\f0\par
}";
        #endregion

        /// <summary>
        /// RTF HTML document code for Google Meet details
        /// </summary>
        private static readonly String rtfHtmlHeader = @"{\rtf1\ansi\ansicpg1252\fromhtml1 \fbidis \deff0{\fonttbl
{\f0\fswiss\fcharset0 Arial;}
{\f1\fmodern Courier New;}
{\f2\fnil\fcharset2 Symbol;}
{\f3\fmodern\fcharset0 Courier New;}}
{\colortbl\red0\green0\blue0;\red5\green99\blue193;}
\uc1\pard\plain\deftab360 \f0\fs24 ";
        private static readonly String rtfHtmlInfo =
        #region HTML document
            @"
{\*\htmltag19 <html xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:w=""urn:schemas-microsoft-com:office:word"" xmlns:m=""http://schemas.microsoft.com/office/2004/12/omml"" xmlns=""http://www.w3.org/TR/REC-html40"">}
{\*\htmltag34 <head>}
{\*\htmltag241 <style>}
{\*\htmltag241 <!--\par /* Font Definitions */\par @font-face\par \tab \{font-family:""Cambria Math"";\par \tab panose-1:2 4 5 3 5 4 6 3 2 4;\par \tab mso-font-charset:0;\par \tab mso-generic-font-family:roman;\par \tab mso-font-pitch:variable;\par \tab mso-font-signature:-536870145 1107305727 0 0 415 0;\}\par @font-face\par \tab \{font-family:Calibri;\par \tab panose-1:2 15 5 2 2 2 4 3 2 4;\par \tab mso-font-charset:0;\par \tab mso-generic-font-family:swiss;\par \tab mso-font-pitch:variable;\par \tab mso-font-signature:-536870145 1073786111 1 0 415 0;\}\par /* Style Definitions */\par p.MsoNormal, li.MsoNormal, div.MsoNormal\par \tab \{mso-style-unhide:no;\par \tab mso-style-qformat:yes;\par \tab mso-style-parent:"""";\par \tab margin:0cm;\par \tab margin-bottom:.0001pt;\par \tab mso-pagination:widow-orphan;\par \tab font-size:11.0pt;\par \tab font-family:""Calibri"",sans-serif;\par \tab mso-ascii-font-family:Calibri;\par \tab mso-fareast-font-family:Calibri;\par \tab mso-hansi-font-family:Calibri;\par \tab mso-bidi-font-family:""Times New Roman"";\par \tab mso-fareast-language:EN-US;\}\par a:link, span.MsoHyperlink\par \tab \{mso-style-noshow:yes;\par \tab mso-style-priority:99;\par \tab color:#0563C1;\par \tab text-decoration:underline;\par \tab text-underline:single;\}\par a:visited, span.MsoHyperlinkFollowed\par \tab \{mso-style-noshow:yes;\par \tab mso-style-priority:99;\par \tab color:#954F72;\par \tab text-decoration:underline;\par \tab text-underline:single;\}\par span.EmailStyle17\par \tab \{mso-style-type:personal-compose;\par \tab mso-style-noshow:yes;\par \tab mso-style-unhide:no;\par \tab font-family:""Calibri"",sans-serif;\par \tab mso-ascii-font-family:Calibri;\par \tab mso-hansi-font-family:Calibri;\par \tab mso-bidi-font-family:Calibri;\par \tab mso-ansi-language:EN;\par \tab mso-no-proof:yes;\}\par .MsoChpDefault\par \tab \{mso-style-type:export-only;\par \tab mso-default-props:yes;\par \tab font-family:""Calibri"",sans-serif;\par \tab mso-ascii-font-family:Calibri;\par \tab mso-fareast-font-family:Calibri;\par \tab mso-hansi-font-family:Calibri;\par \tab mso-bidi-font-family:""Times New Roman"";\par \tab mso-fareast-language:EN-US;\}\par @page WordSection1\par \tab \{size:612.0pt 792.0pt;\par \tab margin:72.0pt 72.0pt 72.0pt 72.0pt;\par \tab mso-header-margin:36.0pt;\par \tab mso-footer-margin:36.0pt;\par \tab mso-paper-source:0;\}\par div.WordSection1\par \tab \{page:WordSection1;\}\par -->}
{\*\htmltag249 </style>}
{\*\htmltag41 </head>}

{\*\htmltag50 <body lang=EN-GB link=""#0563C1"" vlink=""#954F72"" style='tab-interval:36.0pt'>}
{\*\htmltag96 <div class=WordSection1>}\htmlrtf {\htmlrtf0 
{\*\htmltag64 <p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>}\htmlrtf {\htmlrtf0 
{\*\htmltag148 <span lang=EN style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;mso-ansi-language:EN;mso-no-proof:yes'>}\htmlrtf {\htmlrtf0 
{\*\htmltag84 <img width=80 height=30 src=""data:image/png;base64,GMEETLOGO"">}
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 
{\*\htmltag148 <span style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri'>}\htmlrtf {\htmlrtf0 
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 \htmlrtf\par}\htmlrtf0
{\*\htmltag72 </p>}

{\*\htmltag64 <p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>}\htmlrtf {\htmlrtf0 
{\*\htmltag84 <b>}\htmlrtf {\b \htmlrtf0 
{\*\htmltag148 <span style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri'>}\htmlrtf {\htmlrtf0 Google Meet joining information
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 
{\*\htmltag92 </b>}\htmlrtf }\htmlrtf0 
{\*\htmltag72 </p>}

{\*\htmltag64 <p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>}\htmlrtf {\htmlrtf0 
{\*\htmltag148 <span style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri'>}\htmlrtf {\htmlrtf0 
{\*\htmltag84 <a href=""GMEETURL"">}\htmlrtf {\field{\*\fldinst{HYPERLINK ""GMEETURL""}}{\fldrslt\cf1\ul \htmlrtf0 
{\*\htmltag148 <span style='color:blue'>}\htmlrtf {\htmlrtf0 GMEETURL
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 \htmlrtf }\htmlrtf0 \htmlrtf }\htmlrtf0 
{\*\htmltag92 </a>}
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 \htmlrtf\par}\htmlrtf0
{\*\htmltag72 </p>}

{\*\htmltag64 <p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>}\htmlrtf {\htmlrtf0 
{\*\htmltag148 <span style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri'>}\htmlrtf {\htmlrtf0 First time using Meet?
{\*\htmltag84 <a href=""https://gsuite.google.com/learning-center/products/meet/get-started/"">}\htmlrtf {\field{\*\fldinst{HYPERLINK ""https://gsuite.google.com/learning-center/products/meet/get-started/""}}{\fldrslt\cf1\ul \htmlrtf0 
{\*\htmltag84 <b>}\htmlrtf {\b \htmlrtf0 
{\*\htmltag148 <span style='color:#0563C1'>}\htmlrtf {\htmlrtf0 
{\*\htmltag84 &nbsp;}\htmlrtf \'a0\htmlrtf0 Learn more
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 
{\*\htmltag92 </b>}\htmlrtf }\htmlrtf0 \htmlrtf }\htmlrtf0 \htmlrtf }\htmlrtf0 
{\*\htmltag92 </a>}
{\*\htmltag148 <span style='mso-spacerun:yes'>}\htmlrtf {\htmlrtf0 
{\*\htmltag84 &nbsp;}\htmlrtf \'a0\htmlrtf0  
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 \htmlrtf\par}\htmlrtf0
{\*\htmltag72 </p>}

{\*\htmltag64 <p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>}\htmlrtf {\htmlrtf0 
{\*\htmltag148 <span style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri'>}\htmlrtf {\htmlrtf0 
{\*\htmltag244 <o:p>}
{\*\htmltag84 &nbsp;}\htmlrtf \'a0\htmlrtf0 
{\*\htmltag252 </o:p>}
{\*\htmltag156 </span>}\htmlrtf }\htmlrtf0 \htmlrtf\par}\htmlrtf0
{\*\htmltag72 </p>}

{\*\htmltag104 </div>}\htmlrtf }\htmlrtf0 
{\*\htmltag58 </body>}
{\*\htmltag27 </html>}}";
        #endregion

        public static String PlainInfo(String meetingUrl, OlBodyFormat format) {
            if (new OlBodyFormat[] { OlBodyFormat.olFormatPlain, OlBodyFormat.olFormatRichText }.Contains(format))
                return plainInfo.Replace(meetingIdToken, meetingUrl);
            else if (format == OlBodyFormat.olFormatHTML) {
                String hydratedInfo = plainHtmlInfo.Replace(meetingIdToken, meetingUrl);
                return hydratedInfo.Trim();
            } else
                return "";
        }

        public static String RtfInfo(String meetingUrl, Boolean includeHeader = true) {
            String hydratedInfo = rtfInfo.Replace(meetingIdToken, meetingUrl);
            return (includeHeader ? rtfHeader : "") + hydratedInfo;
        }
        public static String RtfHtmlInfo(String meetingUrl, Boolean includeHeader = true) {
            String hydratedInfo = rtfHtmlInfo.Replace(meetingIdToken, meetingUrl);
            hydratedInfo = hydratedInfo.Replace(meetingLogoToken, GMeetLogo.Instance.GMeetLogoBase64);
            return (includeHeader ? rtfHtmlHeader : "") + hydratedInfo;
        }

        private static Regex rgxGmeetUrl = new Regex(@"https:\/\/meet\.google\.com\/[a-z]{3}-[a-z]{4}-[a-z]{3}", RegexOptions.None);
        public static Regex RgxGmeetUrl() {
            return rgxGmeetUrl;
        }

        public static Boolean BodyHasGmeetUrl(this AppointmentItem ai) {
            return rgxGmeetUrl.IsMatch(ai.Body ?? "");
        }

        /// <summary>
        /// Add/update Google Meet information block to Outlook appointment body.
        /// </summary>
        /// <param name="ai">The appointment to update</param>
        /// <param name="gMeetUrl">The URL of the Meeting</param>
        public static void GoogleMeet(this AppointmentItem ai, String gMeetUrl) {
            OlBodyFormat bodyFormat = ai.BodyFormat();
            log.Debug("Body format: " + bodyFormat.ToString());

            if (String.IsNullOrEmpty(gMeetUrl)) {
                CustomProperty.Remove(ref ai, CustomProperty.MetadataId.gMeetUrl);
            } else {
                CustomProperty.Add(ref ai, CustomProperty.MetadataId.gMeetUrl, gMeetUrl);

                if (String.IsNullOrEmpty(ai.Body?.RemoveLineBreaks().Trim())) {
                    log.Debug("Adding GMeet RTF HTML body to Outlook");
                    Calendar.Instance.IOutlook.AddRtfBody(ref ai, RtfHtmlInfo(gMeetUrl));
                } else {
                    if (bodyFormat == OlBodyFormat.olFormatPlain) {
                        if (!ai.BodyHasGmeetUrl()) {
                            log.Debug("Appending GMeet plaintext body to Outlook");
                            ai.Body += "\r\n" + PlainInfo(gMeetUrl, bodyFormat);
                        } else if (String.IsNullOrEmpty(ai.Body?.Replace(PlainInfo(gMeetUrl, bodyFormat), "").RemoveLineBreaks().Trim())) {
                            log.Debug("Replacing GMeet plaintext with RTF body in Outlook");
                            Calendar.Instance.IOutlook.AddRtfBody(ref ai, RtfInfo(gMeetUrl));
                        } else {
                            log.Debug("Updating GMeet plaintext body in Outlook");
                            ai.Body = rgxGmeetUrl.Replace(ai.Body, gMeetUrl);
                        }
                    } else if (bodyFormat == OlBodyFormat.olFormatRichText) {
                        if (!ai.BodyHasGmeetUrl()) {
                            log.Debug("Appending GMeet RTF body to Outlook");
                            String rtfBody = ai.RTFBodyAsString();
                            int injectIdx = rtfBody.LastIndexOf('}');
                            String newRtfBody = rtfBody.Substring(0, injectIdx) + @"\r\n\par\r\n" + RtfInfo(gMeetUrl, false) + rtfBody.Substring(injectIdx + 1);
                            Calendar.Instance.IOutlook.AddRtfBody(ref ai, newRtfBody);
                        } else {
                            log.Debug("Updating GMeet RTF body in Outlook");
                            String newRtfBody = rgxGmeetUrl.Replace(ai.RTFBodyAsString(), gMeetUrl);
                            Calendar.Instance.IOutlook.AddRtfBody(ref ai, newRtfBody);
                        }
                    } else if (bodyFormat == OlBodyFormat.olFormatHTML) {
                        if (!ai.BodyHasGmeetUrl()) {
                            log.Debug("Appending GMeet RTF HTML body to Outlook");
                            String rtfHtmlBody = ai.RTFBodyAsString();
                            int injectIdx = rtfHtmlBody.LastIndexOf(@"{\*\htmltag58 </BODY>}");
                            String newRtfHtmlBody = rtfHtmlBody.Substring(0, injectIdx) + RtfHtmlInfo(gMeetUrl, false);
                            Calendar.Instance.IOutlook.AddRtfBody(ref ai, newRtfHtmlBody);
                        } else {
                            log.Debug("Updating GMeet RTF HTML body in Outlook");
                            String newRtfHtmlBody = rgxGmeetUrl.Replace(ai.RTFBodyAsString(), gMeetUrl);
                            Calendar.Instance.IOutlook.AddRtfBody(ref ai, newRtfHtmlBody);
                        }
                    } else {
                        log.Warn(bodyFormat.ToString() + " is not fully supported. Attempting update of pre-existing GMeet URL.");
                        String newBody = rgxGmeetUrl.Replace(ai.RTFBodyAsString(), gMeetUrl);
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, newBody);
                    }
                }
            }
        }
    }
}
