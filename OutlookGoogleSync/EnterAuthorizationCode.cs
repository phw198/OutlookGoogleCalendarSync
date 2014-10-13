/*
 * Created by SharpDevelop.
 * User: zsianti
 * Date: 14.08.2012
 * Time: 11:58
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of EnterAuthorizationCode.
    /// </summary>
    public partial class EnterAuthorizationCode : Form
    {
        public string authcode = "";
        
        
        public EnterAuthorizationCode()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();
            

        }
        
        void EnterAuthorizationCodeLoad(object sender, EventArgs e)
        {
            
        }
        
        void Button1Click(object sender, EventArgs e)
        {
            authcode = textBox1.Text;
        }


    }
}
