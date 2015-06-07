using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices; 


namespace Speaker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
          
        }
        

        StreamReader CommandsReader = new StreamReader(@"C:\users\" + Environment.UserName.ToString()+@"\documents\commands.txt");
        SpeechSynthesizer sSynth = new SpeechSynthesizer();
        PromptBuilder pBuilder = new PromptBuilder();
        SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();
        string name = "Microsoft Zira Desktop"; 

        Point lastPosition; 
        
         
        private void Form1_Load(object sender, EventArgs e)
        {

            initializeSpeach();          
            
        }

       
        GrammarBuilder gbuilder = new GrammarBuilder();
         public void initializeSpeach()
            {


            Choices sList = new Choices();

            //Add the words
        


            try
            {
                // the command txt file

                gbuilder.Append(new Choices(System.IO.File.ReadAllLines("commands.txt")));

            }
            catch { MessageBox.Show("The 'Commands' file must not contain empty lines.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); pr.StartInfo.FileName = @"C:\users\" + Environment.UserName.ToString()+@"\documents\commands.txt"; pr.Start(); Application.Exit(); return; }
            Grammar gr = new Grammar(gbuilder);
            try
            {
                sRecognize.UnloadAllGrammars();
                sRecognize.RecognizeAsyncCancel();
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;
                
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);
               
                


            }

            catch
            {
                MessageBox.Show("Grammar Builder Error"); 
                return;
            }
           
        }

     
     bool start = false;
         
     Process pr = new Process(); 
     public void lockComputer()
     {

         System.Diagnostics.Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
         return; 
     }

         public void speakText(string textSpeak)
        {
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            pBuilder.ClearContent();
            pBuilder.AppendText(textSpeak.ToString());
            sSynth.SelectVoice(name);
            sSynth.SpeakAsync(pBuilder);
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            sRecognize.RecognizeAsync(RecognizeMode.Multiple);
        }
         bool exitCondition = false;
        private void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
     {
         

            //THIS IS WHERE YOU ADD COMMANDS

            //Take note: GREEN TEXT IS CODE THAT IS STILL IN DEVELOPMENT. 
            //YOU MAY CHOOSE TO USE IT OR DELETE IT. THIS DOCUMENT WILL BE 
            //UPDATED EVERY TIME I MAKE A MAJOR UPDATE.


         //if(exitCondition)
         //{
         //    Thread.Sleep(100);
         //    if (e.Result.Text == "yes")
         //    {
         //        sSynth.SpeakAsyncCancelAll();
         //        sRecognize.RecognizeAsyncCancel();
         //        Application.Exit();

         //        return;
         //    }
         //    else { exitCondition = false;  speakText("Exit Cancelled"); return; }
         //}
         switch (e.Result.Text)
         {
             case "name":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("iam voice recognitio application");
                 break;

             case "old":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("iam still baby");
                 break;

             case "hello":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Hello uncle eysa");
                 break;

             case "invisable":
                 if (ActiveForm.Visible == true)
                 {
                    listBox2.Items.Add(e.Result.Text.ToString());
                    ActiveForm.ShowInTaskbar = false;  ActiveForm.Hide(); 
                     
                                   
                     speakText("I am now invisible. You can access me by clicking on the icon down here in the tray.");
                     break;
                 }
                 else
                 {
                      
                     break; 
                 }
             case "salam":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 //speakText("Hello, " + Environment.UserName.ToString());
                 speakText("waalaykom el-salam warahmato allah wabarakato");
                 break;

             case "exit":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Are you sure you want to exit?");
                 Thread.Sleep(3000);
                 Application.Exit();
                 //MessageBox.Show("are you sure");

                 //exitCondition = true; 
                 break;

             case "no":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("welcome ya sleeem ya zalooooot?");
                 

                 exitCondition = true;
                 break;

             case "thank you":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("ala eih yasta eysa habeeb alby");
                 break;
             case"change voice":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 foreach (InstalledVoice voice in sSynth.GetInstalledVoices())
                 {
                     bool first = true;
                     StreamReader sr = new StreamReader(@"C:\Temp\Speakvoice.txt");
                     string name = sr.ReadLine();
                     sr.Close();
                     VoiceInfo vi = voice.VoiceInfo;
                     //comboBox1.Items.Add(vi.Name.ToString());
                     if (first)
                     {
                         first = false;
                  //       comboBox1.Text = name;
                     }

                 }
               //  comboBox1.Show();
                 break; 
             case "stop talking":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 cancelSpeech();
                 break; 
             case "lock computer":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 lockComputer();
                 break;

              

             case "be quiet":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 cancelSpeech();
                 sSynth.SpeakAsyncCancelAll();
                 break;

             case "chrome":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("One moment yasta.");
                 //Process pr = new Process();
                 //pr.StartInfo.FileName = "http://www.google.com/";
                 //pr.Start();
                 Process.Start("chrome.exe", "");

                 break; 

             case "stop listening":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Ok.");
              //   btnStop.Visible = true;
                 sRecognize.RecognizeAsyncCancel();
                 sRecognize.RecognizeAsyncStop();
                 break;

             case "calculator":
                 Process p2 = new Process();
                 p2 = System.Diagnostics.Process.Start("calc.exe");
                 break; 


             case "show commands":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 bool working = true; 
                 while (working)
                 {
                     try
                     {
                         listBox1.Show();
                      listBox1.Items.Add(CommandsReader.ReadLine());
                     }
                     catch { working = false; break; }

                 }
                 break;
             case "maximize":
                 
                 try
                 {
                     ActiveForm.WindowState = FormWindowState.Maximized;
                     listBox2.Items.Add(e.Result.Text.ToString());
                     break;
                 }
                 catch
                 {
                     listBox2.Items.Add(e.Result.Text.ToString()+ "[FAILED]") ;
                     break;
                 }
             case "if one":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 bool worrking = true; 
                 while (worrking)
                 {
                     try
                     {
                         listBox1.Show();
                      listBox1.Items.Add(CommandsReader.ReadLine());
                     }
                     catch { worrking = false; break; }

                 }

                 break;
             case "eject drive":
                 OpenCloseCD();
                 listBox2.Items.Add(e.Result.Text.ToString());
                 
                 break; 
             case "hide":
                 break;

             case "what about you":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("iam an artificial intelegence application");
                 break;

             case "your friend":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("it is  muhammad eysa");
                 break;

             case "my facebook account":
                 Process.Start("chrome.exe", "https://www.facebook.com/eng.eissa.39");
                 break;
         }


       
            
        
         
         
     }
        
    
       
        private void button3_Click(object sender, EventArgs e)
        {
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RequestRecognizerUpdate();
            sRecognize.UnloadAllGrammars();
            
             
        }
        
        public void OpenCloseCD()
        {
            EjectMedia.Eject(@"\\.\D:");

            
            
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            pBuilder.ClearContent(); 
            sRecognize.RecognizeAsyncCancel(); 
        }

        private void addWordsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show(); 
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                lastPosition = e.Location;
            }
            else { return; } 
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                this.Left += e.X - lastPosition.X;
                this.Top += e.Y - lastPosition.Y;
            }
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            sSynth.SpeakAsyncCancelAll(); 
        }







        internal void cancelSpeech()
        {
            sSynth.SpeakAsyncCancelAll(); 
        }

        

        private void btnStop_Click(object sender, EventArgs e)
        {
            sRecognize.RecognizeAsync(RecognizeMode.Multiple);
       //     btnStop.Visible = false; 
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        //    name = comboBox1.Text.ToString();
        //    sSynth.SelectVoice(comboBox1.Text);
            
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            EjectMedia.Eject(@"\\.\D:");
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }
    }
}
