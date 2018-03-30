﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using InventorApprentice;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Archie
{
    public partial class Form1 : Form
    {
        public Inventor.Application _invApp;
        public Boolean _started = false;
        private InventorApprentice.ApprenticeServerComponent m_oserver;
        static SerialPort serial1 = new SerialPort("COM1", 9600);
        private MoveAndRotate_1.InventorCamera movrot = null;

        Inventor.ApprenticeServerDocument oDoc;
        Inventor.ApprenticeServerComponent oApp;

        protected override void WndProc(ref Message m)
        {
            const int WM_NCLBUTTONDOWN = 161;
            const int WM_SYSCOMMAND = 274;
            const int HTCAPTION = 2;
            const int SC_MOVE = 61456;

            if ((m.Msg == WM_SYSCOMMAND) && (m.WParam.ToInt32() == SC_MOVE))
            {
                return;
            }

            if ((m.Msg == WM_NCLBUTTONDOWN) && (m.WParam.ToInt32() == HTCAPTION))
            {
                return;
            }

            base.WndProc(ref m);
        }

        public Form1()
        {
            try
            {
                InitializeComponent();
                m_oserver = new InventorApprentice.ApprenticeServerComponent();
                movrot = new MoveAndRotate_1.InventorCamera();
                movrot.GetCurrentCamera();
                // AddInventorPath();
                try
                {
                    _invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                    _invApp.UserInterfaceManager.RibbonDockingState = RibbonDockingStateEnum.kDockToLeft;

                }
                catch (Exception ex)
                {
                    try
                    {

                        Type invAppType = Type.GetTypeFromProgID("Inventor.Application");
                        _invApp = (Inventor.Application)Activator.CreateInstance(invAppType);
                        _invApp.Visible = true;
                        _started = true;
                    }
                    catch (Exception ex2)
                    {

                        MessageBox.Show("Wait, reopen your application", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _invApp.UserInterfaceManager.RibbonDockingState = RibbonDockingStateEnum.kDockToLeft;
                    }

                }
            }
            catch
            {
                MessageBox.Show(this,
                  "Failed to create an instance of Apprentice server.",
                  "CSharpFileDisplaySample",
                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void serial1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string[] input = serial1.ReadLine().Split(',');
            int num = int.Parse(input[0]);
            int x = int.Parse(input[1]);
            int y = int.Parse(input[2]);
            if (num == 5)
            {
                this.BeginInvoke(new LineReceivedEvent(LineReceived), x, y);
            }
            else if (num == 2)
            {
                this.BeginInvoke(new LineReceivedEvent(ExtrudeSketch), x, y);
            }
            else if (num == 3)
            {
                this.BeginInvoke(new LineReceivedEvent(ExtrudeOnly), x, y);
            }
            else if (num == 0)
            {
                this.BeginInvoke(new LineReceivedEvent(ClickHand), x, y);
            }
        }

        private delegate void LineReceivedEvent(int x, int y);

        private void LineReceived(int x, int y)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
            DoClickMouse(0x4);
        }

        private void ClickHand(int x, int y)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
            DoClickMouse(0x2);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

            if (serial1.IsOpen == true)
            {
                try
                {
                    serial1.PortName = "COM1";
                    serial1.BaudRate = 9600;
                    serial1.DataReceived += new SerialDataReceivedEventHandler(serial1_DataReceived);
                    textBox1.Text = "Connected";
                    pictureBox1.Enabled = false;
                    label1.Enabled = false;
                    pictureBox2.Enabled = true;
                    label2.Enabled = true;
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("COM already used. Disconnect your previous communication", "Dual Port Issue",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
            else if (serial1.IsOpen == false)
            {
                try
                {
                    serial1.PortName = "COM1";
                    serial1.BaudRate = 9600;
                    serial1.Open();
                    serial1.DataReceived += new SerialDataReceivedEventHandler(serial1_DataReceived);
                    textBox1.Text = "Connected";
                    pictureBox1.Enabled = false;
                    label1.Enabled = false;
                    pictureBox2.Enabled = true;
                    label2.Enabled = true;
                }

                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("COM already used. Disconnect your previous communication", "Dual Port Issue",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            serial1.Close();
            pictureBox2.Enabled = false;
            label2.Enabled = false;
            pictureBox1.Enabled = true;
            label1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(20, 35);

            DoClickMouse(0x2); 
            DoClickMouse(0x4);
        }

        static void DoClickMouse(int mouseButton)
        {
            var input = new INPUT()
            {
                dwType = 0, 
                mi = new MOUSEINPUT() { dwFlags = mouseButton }
            };

            if (SendInput(1, input, Marshal.SizeOf(input)) == 0)
            {
                throw new Exception();
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            int dx;
            int dy;
            int mouseData;
            public int dwFlags;
            int time;
            IntPtr dwExtraInfo;
        }
        struct INPUT
        {
            public uint dwType;
            public MOUSEINPUT mi;
        }
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint cInputs, INPUT input, int size);

        private void ExtrudeSketch(int x, int y)
        {
            //PartDocument oPartDoc = default(PartDocument);
            //oPartDoc = (PartDocument)_invApp.Documents.Add(Inventor.DocumentTypeEnum.kPartDocumentObject,_invApp.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kPartDocumentObject));
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            PlanarSketch oSketch = default(PlanarSketch);
            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[1]);
            oSketch.SketchLines.AddAsTwoPointRectangle(_invApp.TransientGeometry.CreatePoint2d(-x, -y),_invApp.TransientGeometry.CreatePoint2d(x, y));

            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();

            ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
            oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
            oExtrudeDef.SetDistanceExtent(5, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
        }

        private void SketchOnly(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            PlanarSketch oSketch = default(PlanarSketch);
            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[1]);
            oSketch.SketchLines.AddAsTwoPointRectangle(_invApp.TransientGeometry.CreatePoint2d(-x, -y), _invApp.TransientGeometry.CreatePoint2d(x, y));
        }

        private void ExtrudeOnly(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures["Extrusion1"];
            oExtrude.SetDistanceExtent(x, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //ExtrudeOnly();
            movrot.ChangeView(4, 5, 7, apply:true);
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            _invApp.UserInterfaceManager.RibbonDockingState = RibbonDockingStateEnum.kDockToTop;
        }
    }
}