// This code is distributed under MIT license. 
// Copyright (c) 2015 George Mamaladze
// See license.txt or http://opensource.org/licenses/mit-license.php

using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;
using System.Windows.Automation;
using System.Threading;
using Gma.System.MouseKeyHook.Implementation;
using System.Text;
using System.Runtime.InteropServices;

namespace Demo
{
    public partial class Main : Form
    {
        private IKeyboardMouseEvents m_Events;

        public Main()
        {
  
            InitializeComponent();
            radioGlobal.Checked = true;
            SubscribeGlobal();
            FormClosing += Main_Closing;

        }

        private void Main_Closing(object sender, CancelEventArgs e)
        {
            Unsubscribe();
        }

        private void SubscribeApplication()
        {
            Unsubscribe();
            Subscribe(Hook.AppEvents());
        }

        private void SubscribeGlobal()
        {
            Unsubscribe();
            Subscribe(Hook.GlobalEvents());
        }

        /*
        // I want to subscribe and handle all events in less code, is that possible?
        // 
        private void Subscribe(IKeyboardMouseEvents events)
        {
            
            
            // Subscribe to 
            for ev in events:
                ev += Foo;
        }

        // Won't really work because args is different. But we can at 
        // least call the same print function
        private void Foo(object sender, args e)
        {
            // Which event occurred?!?!

            // Print out event name. Make it human readable, but compress
            // when we are done. 
            event.ToString()
        }
        */

        private void Subscribe(IKeyboardMouseEvents events)
        {
            m_Events = events;
            m_Events.KeyDown += OnKeyDown;
            m_Events.KeyUp += OnKeyUp;
            m_Events.KeyPress += HookManager_KeyPress;

            m_Events.MouseUp += OnMouseUp;
            m_Events.MouseClick += OnMouseClick;
            m_Events.MouseDoubleClick += OnMouseDoubleClick;

            m_Events.MouseMove += HookManager_MouseMove;

            m_Events.MouseDragStarted += OnMouseDragStarted;
            m_Events.MouseDragFinished += OnMouseDragFinished;

            if (checkBoxSupressMouseWheel.Checked)
                m_Events.MouseWheelExt += HookManager_MouseWheelExt;
            else
                m_Events.MouseWheel += HookManager_MouseWheel;

            if (checkBoxSuppressMouse.Checked)
                m_Events.MouseDownExt += HookManager_Supress;
            else
                m_Events.MouseDown += OnMouseDown;
            
        }

        private void Unsubscribe()
        {
            if (m_Events == null) return;
            m_Events.KeyDown -= OnKeyDown;
            m_Events.KeyUp -= OnKeyUp;
            m_Events.KeyPress -= HookManager_KeyPress;

            m_Events.MouseUp -= OnMouseUp;
            m_Events.MouseClick -= OnMouseClick;
            m_Events.MouseDoubleClick -= OnMouseDoubleClick;

            m_Events.MouseMove -= HookManager_MouseMove;

            m_Events.MouseDragStarted -= OnMouseDragStarted;
            m_Events.MouseDragFinished -= OnMouseDragFinished;

            if (checkBoxSupressMouseWheel.Checked)
                m_Events.MouseWheelExt -= HookManager_MouseWheelExt;
            else
                m_Events.MouseWheel -= HookManager_MouseWheel;

            if (checkBoxSuppressMouse.Checked)
                m_Events.MouseDownExt -= HookManager_Supress;
            else
                m_Events.MouseDown -= OnMouseDown;

            m_Events.Dispose();
            m_Events = null;
        }

        private void HookManager_Supress(object sender, MouseEventExtArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                Log(string.Format("MouseDown \t\t {0}\n", e.Button));
                return;
            }

            Log(string.Format("MouseDown \t\t {0} Suppressed\n", e.Button));
            e.Handled = true;
        }
        /// <summary>
        /// Need to do some things special with keyboard presses. Particularly
        /// with Alt+Tab, we ignore all key down until key up (the final window switch)
        /// 
        /// 
        /// Log the active window that the event took place on. Mouse clicks should 
        /// be logged as well so we'll be keyboard pressing in the correct spot.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            Log(string.Format("KeyDown  \t\t {0}\n", e.KeyCode));
        }

        private void OnKeyUp(object sender, KeyEventArgs e)
        {
            Log(string.Format("KeyUp  \t\t {0}\n", e.KeyCode));
        }

        private void HookManager_KeyPress(object sender, KeyPressEventArgs e)
        {
            Log(string.Format("KeyPress \t\t {0}\n", e.KeyChar));
        }

        private void HookManager_MouseMove(object sender, MouseEventArgs e)
        {
            labelMousePosition.Text = string.Format("x={0:0000}; y={1:0000}", e.X, e.Y);
        }

        /// <summary>
        /// Not needed for logging at this point. But planning on adding very 
        /// soon. Really helpful...
        /// 
        /// </summary>
        private void ShowAllButtons()
        {
            // Show unique keyboard shortcuts, but then only save x and y for each
            // so you can just do ElementFromPoint(x,y)
            TreeWalker walker = TreeWalker.ControlViewWalker;
            int MAX_DEPTH = 100;
            AutomationElement[] elements = new AutomationElement[MAX_DEPTH];
            int i = 0;
            elements[i] = AutomationElement.RootElement; //Use GetActiveWindow()??
            i++;
            AutomationElement ae;
            while (i>0)
            {
                ae = walker.GetNextSibling(elements[i-1]);
                if (ae == null)
                {
                    //Add appropriate label and x,y coordinate of parent to some list somewhere
                    // Check if it's clickable too

                    // Check if it's something we *can* click (it's "enabled"). By using the controlViewWalker
                    /*
# we are only getting things that are "controllable"/clickable??

# TODO: Limit by IsInvokePatternAvailable or LegacyIAccessible.State == 'Linked' or 'normal'
# Also, add to an "index" in case we want to search by keyword later


# TODO: Limit by IsInvokePatternAvailable or LegacyIAccessible.State == 'Linked' or 'normal'
# Also, add to an "index" in case we want to search by keyword later
                    if uiElement.CurrentIsOffscreen == False: # and uiElement.CurrentIsEnabled == True:
                # Thinking about later stuff
                #if uiElement.CurrentLocalizedControlType == 'menu item':
                    
                
                tree.append(uiElement)
                rect = (uiElement.CurrentBoundingRectangle.left,
                                         uiElement.CurrentBoundingRectangle.top,
                                         uiElement.CurrentBoundingRectangle.right,
                                         uiElement.CurrentBoundingRectangle.bottom)
                        */
                    i--;
                    continue;
                }
                // Go down a level
                elements[i] = ae;
                i++;
                continue;
            }
        }
        private AutomationElement ElementFromPoint(int x, int y)
        {
            //Or pass in parameters, not sure ... FromPoint(e.X, e.Y);
            // Convert mouse position from System.Drawing.Point to System.Windows.Point.
            System.Windows.Point point = new System.Windows.Point(x,y);
            AutomationElement element = AutomationElement.FromPoint(point);
            ToolTip tt = new ToolTip();
            IWin32Window win = this;
            tt.ShowAlways = true;
            //Should be using Cursor.Position
            tt.Show("foo", win, 300,300,1000);
            return element;
        }

        private string GetXPathFromElement(AutomationElement ae)
        {
            string path = "";
            TreeWalker walker = TreeWalker.ControlViewWalker;

            while (true)
            {
                try
                {
                    if (ae == null)
                    {
                        break;
                    }
                    path = ae.Current.Name + "/" + path;
                    if (ae == AutomationElement.RootElement) break;
                    ae = walker.GetParent(ae);
                } catch (ElementNotAvailableException e)
                {
                    //Window probably closed. 
                    //TODO: figure out what to do here
                    break;
                }
                
            }
            return path;
        }

        private AutomationElement GetElementFromXPath(string path)
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement ae = AutomationElement.RootElement;
            char[] separatingChars = { '/' };

            if (path.Equals(""))
            {
                return null;
            }
            //Remove empty entries...w00t
            // Remove initial blank
            path = path.Remove(0, 1);
            string[] words = path.Split(separatingChars);
            string currentName;

            TreeScope ts = TreeScope.Children;
            int i = 0;
            foreach (string s in words)
            {
                if (ae == null || ae.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty).ToString().ToLower().Equals("false"))
                {
                    return null;
                }
                currentName = ae.Current.Name;
                if (s.Equals(""))
                {
                    // String in path is nothing. 
                    // So delay searching until we have something for sure
                    ts = TreeScope.Descendants;
                    continue;
                }
                else
                {
                    ae = ae.FindFirst(ts, new PropertyCondition(AutomationElement.NameProperty, s));
                    //Set default back to searching just children
                    ts = TreeScope.Children;
                }
                System.Diagnostics.Debug.Write(string.Format("Next:{0}, {1}\n", s, s.Equals("")));
                i++;
            }
            
            return ae;
        }

        private void backgroundWorker1_DoWork(
                object sender,
                DoWorkEventArgs e)
        {
            if (e.Argument == null)
            {
                System.Diagnostics.Debug.Write(string.Format("BOO HISS NULL\n"));
                e.Result = null;
            }
            else
            {
                e.Result = GetXPathFromElement(e.Argument as AutomationElement);
                if (e.Result != null)
                {
                    // Prove that we can re-navigate the tree and locate the element
                    // again.
                    AutomationElement foo = GetElementFromXPath((string)e.Result);
                    if (foo == null)
                    {
                        // The object searched for by xpath doesn't exist anymore.
                        // What to do, what to do..
                        //TODO
                    }
                    else
                    {
                        e.Result += "\n Object Name: " + foo.Current.Name;
                    }
                }
            }
            return;
        }

        private void backgroundWorker1_RunWorkerCompleted(
                object sender,
                RunWorkerCompletedEventArgs e)
        {
            Log(string.Format("Parents \t\t {0}\n", e.Result));
        }

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            Log(string.Format("MouseDown \t\t {0}\n", e.Button));
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            Log(string.Format("MouseUp \t\t {0}\n", e.Button));
        }

        private void OnMouseClick(object sender, MouseEventArgs e)
        {
            AutomationElement ae;
            ae = ElementFromPoint(Cursor.Position.X, Cursor.Position.Y);
            AutomationProperty enabledProp = AutomationElement.IsEnabledProperty;
            if (ae.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty).ToString().ToLower().Equals("true"))
            {
                if (!this.backgroundWorker1.IsBusy)
                    this.backgroundWorker1.RunWorkerAsync(ae);
                else
                    MessageBox.Show("Can't run the worker twice!");
                
            } else
            {
                //Window died probably.
                ae = null;
            }
            
            Log(string.Format("MouseClick \t\t {0}\n", e.Button));
        }

        private void OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            Log(string.Format("MouseDoubleClick \t\t {0}\n", e.Button));
        }

        private void OnMouseDragStarted(object sender, MouseEventArgs e)
        {
            Log("MouseDragStarted\n");
        }

        private void OnMouseDragFinished(object sender, MouseEventArgs e)
        {
            Log("MouseDragFinished\n");
        }

        private void HookManager_MouseWheel(object sender, MouseEventArgs e)
        {
            labelWheel.Text = string.Format("Wheel={0:000}", e.Delta);
        }
        
        private void HookManager_MouseWheelExt(object sender, MouseEventExtArgs e)
        {
            labelWheel.Text = string.Format("Wheel={0:000}", e.Delta);
            Log("Mouse Wheel Move Suppressed.\n");
            e.Handled = true;
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        private string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        /// <summary>
        /// Format of log entry should be:
        /// Timestamp;; type of press;; object clicked or keys pressed;;
        /// 
        /// If we run into issues later with parsing, use non-ascii characters
        /// for non-keyboard events?
        /// </summary>
        /// <param name="text"></param>
        private void Log(string text)
        {
            if (IsDisposed) return;
            string text2 = DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds +
                        "Active Window: " + GetActiveWindowTitle() + text;
            using (StreamWriter w = File.AppendText("log.txt"))
            {
                //TODO: We should also log active window of keyboard presses to make 
                // things a little simpler in reproducing.
                // We are logging the xPath of clicks, which includes the parent window
                //http://stackoverflow.com/questions/115868/how-do-i-get-the-title-of-the-current-active-window-using-c




                // Let's use unix time, since we don't care about actual time, but more relative 
                // time without dealing with all the time zone and DST gunk.
                w.WriteLine(text2);
                /*"{0} Active Window: {1} {2}", (DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds,
                //DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                //DateTime.Now.ToLongTimeString(),DateTime.Now.ToLongDateString(),
                GetActiveWindowTitle(),
            text);
            */
            }
            
            textBoxLog.AppendText(text2);
            textBoxLog.ScrollToCaret();
        }

        private void radioApplication_CheckedChanged(object sender, EventArgs e)
        {
            if (((RadioButton) sender).Checked) SubscribeApplication();
        }

        private void radioGlobal_CheckedChanged(object sender, EventArgs e)
        {
            if (((RadioButton) sender).Checked) SubscribeGlobal();
        }

        private void radioNone_CheckedChanged(object sender, EventArgs e)
        {
            if (((RadioButton) sender).Checked) Unsubscribe();
        }

        private void checkBoxSupressMouseWheel_CheckedChanged(object sender, EventArgs e)
        {
            if (m_Events == null) return;

            if (((CheckBox) sender).Checked)
            {
                m_Events.MouseWheel -= HookManager_MouseWheel;
                m_Events.MouseWheelExt += HookManager_MouseWheelExt;
            }
            else
            {
                m_Events.MouseWheelExt -= HookManager_MouseWheelExt;
                m_Events.MouseWheel += HookManager_MouseWheel;
            }
        }

        private void checkBoxSuppressMouse_CheckedChanged(object sender, EventArgs e)
        {
            if (m_Events == null) return;

            if (((CheckBox)sender).Checked)
            {
                m_Events.MouseDown -= OnMouseDown;
                m_Events.MouseDownExt += HookManager_Supress;
            }
            else
            {
                m_Events.MouseDownExt -= HookManager_Supress;
                m_Events.MouseDown += OnMouseDown;
            }
        }

        private void clearLog_Click(object sender, EventArgs e)
        {
            textBoxLog.Clear();
        }
    }
}