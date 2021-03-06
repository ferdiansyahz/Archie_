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
using System.Threading;

namespace Archie
{
    public partial class Form1 : Form
    {
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int)m.Result == 0x1)
                        m.Result = (IntPtr)0x2;
                    return;
            }

            base.WndProc(ref m);
        }

        public Inventor.Application _invApp;
        public Boolean _started = false;
        private InventorApprentice.ApprenticeServerComponent m_oserver;
        static SerialPort serial1 = new SerialPort("COM1", 9600);
        private MoveAndRotate_1.InventorCamera movrot = null;
        public PartDocument oPartDoc = default(PartDocument);

        Inventor.ApprenticeServerDocument oDoc;
        Inventor.ApprenticeServerComponent oApp;
        Stopwatch stopwatch = new Stopwatch();
        
        /*protected override void WndProc(ref Message m)
        {
            const int wM_NCLBUTTONDOWN = 161;
            const int wM_SYSCOMMAND = 274;
            const int HTCAPTION = 2;
            const int SC_MOVE = 61456;
            if ((m.Msg == wM_SYSCOMMAND) && (m.WParam.ToInt32() == SC_MOVE))
            {
                return;
            }

            if ((m.Msg == wM_NCLBUTTONDOWN) && (m.WParam.ToInt32() == HTCAPTION))
            {
                return;
            }

            base.WndProc(ref m);
        }*/

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
                    //oPartDoc = (PartDocument)_invApp.Documents.Add(Inventor.DocumentTypeEnum.kPartDocumentObject, _invApp.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kPartDocumentObject));
                    PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
                    PartComponentDefinition oCompDef = default(PartComponentDefinition);
                    oCompDef = oPartDoc.ComponentDefinition;
                    foreach (Inventor.PartFeature oFeat in oCompDef.Features)
                    {
                        oFeat.Delete();
                    }
                    movrot = new MoveAndRotate_1.InventorCamera();
                    movrot.GetCurrentCamera();
                }
                catch (Exception ex)
                {
                    try
                    {

                        Type invAppType = Type.GetTypeFromProgID("Inventor.Application");
                        _invApp = (Inventor.Application)Activator.CreateInstance(invAppType);
                        _invApp.Visible = true;
                        _started = true;
                        movrot = new MoveAndRotate_1.InventorCamera();
                        movrot.GetCurrentCamera();
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
            try
            {
                string[] input = serial1.ReadLine().Split(',');
                
                int num = int.Parse(input[0]);
                int x = int.Parse(input[1]);
                int y = int.Parse(input[2]);
                //this.BeginInvoke(new LineReceivedEvent(SaveCSV), x, y);

                if (num == 5)
                {
                    //this.BeginInvoke(new LineReceivedEvent(LineReceived), x, y);
                    //this.BeginInvoke(new LineReceivedEvent(CircleSketch), x, y);
                    try
                    {
                        movrot.ReturnHome();
                        this.BeginInvoke(new LineReceivedEvent(Feature1), x, y);
                        
                    }
                    catch
                    {

                    }

                }
                else if (num == 2)
                {
                    //this.BeginInvoke(new LineReceivedEvent(CircleSketch), x, y);
                    /*PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
                    PartComponentDefinition oCompDef = default(PartComponentDefinition);
                    oCompDef = oPartDoc.ComponentDefinition;
                    foreach (Inventor.PartFeature oFeat in oCompDef.Features)
                    {
                        oFeat.Delete();
                    }*/
                    try
                    {
                        movrot.RotateCube("LEFT");
                        this.BeginInvoke(new LineReceivedEvent(Feature2), x, y);
                        
                    }
                    catch
                    {

                    }
                }
                else if (num == 3)
                {
                    //this.BeginInvoke(new LineReceivedEvent(ExtrudeDoub), x, y);
                    try
                    {
                        this.BeginInvoke(new LineReceivedEvent(Feature3), x, y);
                  
                    }
                    catch
                    {

                    }
                }
                else if (num == 1)
                {
                    //this.BeginInvoke(new LineReceivedEvent(ClickHand), x, y);
                    //this.BeginInvoke(new LineReceivedEvent(SaveCSV), x, y);  
                 
                }
                else if (num == 4)
                {
                    //this.BeginInvoke(new LineReceivedEvent(moving), x, y);
                    //this.BeginInvoke(new LineReceivedEvent(ExtrudeDoub), x, y);
                    try
                    {
                        this.BeginInvoke(new LineReceivedEvent(Feature4), x, y);
                       
                    }
                    catch
                    {

                    }
                }
                else if (num == 0)
                {
                    //this.BeginInvoke(new LineReceivedEvent(rotateG), x, y);
                    try
                    {
                        this.BeginInvoke(new LineReceivedEvent(Feature5), x, y);
                        movrot.RotateCube("LEFT");
                        movrot.ReturnHome();
                    }
                    catch
                    {

                    }
                }
            
                
            }
            catch
            {
                return;
            }
        }

        private delegate void LineReceivedEvent(int x, int y);

        private void LineReceived(int x, int y)
        {
            //System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
            //DoClickMouse(0x4);
            SaveCSV_SerialRead();
            SaveCSV(x, y);
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

        private void ExtrudeOnly(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures["Extrusion1"];
            oExtrude.SetDistanceExtent(x, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);    
        }

        private void ExtrudeDoub(int x, int y)
        {
            try
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;

                ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures[1];
                oExtrude.SetDistanceExtent(x, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

                Random r = new Random();
                int rand = r.Next(1, 20);

                if (rand > 18)
                {
                    movrot.ChangeView(1, 0, 0, apply: false);
                }

            }

            catch (System.ArgumentException)
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;

                PlanarSketch oSketch = default(PlanarSketch);
                oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[1]);
                oSketch.SketchLines.AddAsTwoPointRectangle(_invApp.TransientGeometry.CreatePoint2d(-x, -y), _invApp.TransientGeometry.CreatePoint2d(x, y));

                Profile oProfile = default(Profile);
                oProfile = oSketch.Profiles.AddForSolid();

                ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
                oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
                oExtrudeDef.SetDistanceExtent(x, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

                ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
            }
        }
        
        private void moving(int x, int y)
        {
            if(x > 5 && y > 6) //Kuadran 2
            {
                movrot.TranslateView(-1, -1);
            }

            else if (x < 5 && y > 6) //kuadran 1
            {
                movrot.TranslateView(1, -1);
            }

            else if (x < 5 && y < 6) //kuadran 3
            {
                movrot.TranslateView(1, 1);
            }

            else if (x > 5 && y < 6) //kuadran 3
            {
                movrot.TranslateView(-1, 1);
            }
        }

        private void rotateG(int x, int y)
        {
            if (x > 5 && y > 6) //Kuadran 2
            {
                movrot.ChangeView(1, 0, 0, apply: false);
            }

            else if (x < 5 && y > 6) //kuadran 1
            {
                movrot.ChangeView(-1, 0, 0, apply: false);
            }

            else if (x < 5 && y < 6) //kuadran 3
            {
                movrot.ChangeView(0, 1, 0, apply: false);
            }

            else if (x > 5 && y < 6) //kuadran 3
            {
                movrot.ChangeView(0, -1, 0, apply: false);
            }
        }

        private void CircleSketch(int x, int y)
        {
            try
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;

                PlanarSketch oSketch = default(PlanarSketch);
                oSketch = oCompDef.Sketches[1];
                oSketch.SketchCircles[1].Radius = x;
                oSketch.SketchCircles[2].Radius = x - 2;

                Profile oProfile = default(Profile);
                oProfile = oSketch.Profiles.AddForSolid();

                ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures[1];
                oExtrude.Definition.Profile = oProfile;
                oExtrude.Definition.SetDistanceExtent(y, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                /*ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures[1];     
                oExtrude.*/
                Random r = new Random();
                int rand = r.Next(1, 20);

                if (rand > 17)
                {
                    int ner = r.Next(1, 3);
                    movrot.ChangeView(ner, 0, 0, apply: false);
                }

                if (x > 450 && y > 300)
                {
                    oCompDef = oPartDoc.ComponentDefinition;
                    foreach (Inventor.PartFeature oFeat in oCompDef.Features)
                    {
                        oFeat.Delete();
                    }
                    foreach (Inventor.Sketch oSket in oCompDef.Sketches)
                    {
                        oSket.Delete();
                    }
                }

                if(rand == 15)
                {                  
                    oCompDef = oPartDoc.ComponentDefinition;

                    oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[1]);
                    oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), x);
                    oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), x - 2);
                    
                    oProfile = oSketch.Profiles.AddForSolid();

                    ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
                    oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
                    oExtrudeDef.SetDistanceExtent(y, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

                    oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
                }

            }
            catch
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;

                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;

                PlanarSketch oSketch = default(PlanarSketch);
                oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[1]);
                oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), x);
                oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), x - 2);

                Profile oProfile = default(Profile);
                oProfile = oSketch.Profiles.AddForSolid();

                ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
                oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
                oExtrudeDef.SetDistanceExtent(y, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

                ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
            }
        }

        private void Feature1(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
            oPartDoc.UnitsOfMeasure.LengthUnits = Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            PlanarSketch oSketch = default(PlanarSketch);
            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 6);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 8);

            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();

            ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
            oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
            oExtrudeDef.SetDistanceExtent(14, Inventor.PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
        }

        private void Feature2(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
            oPartDoc.UnitsOfMeasure.LengthUnits = Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            PlanarSketch oSketch = default(PlanarSketch);
            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 15);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 17.5);

            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();

            ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
            oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
            oExtrudeDef.SetDistanceExtent(14, Inventor.PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
        }

        private void Feature3(int x, int y)
        {
            PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
            oPartDoc.UnitsOfMeasure.LengthUnits = Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
            PartComponentDefinition oCompDef = default(PartComponentDefinition);
            oCompDef = oPartDoc.ComponentDefinition;

            PlanarSketch oSketch = default(PlanarSketch);
            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 6);
            oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 17.5);

            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();

            ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
            oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kJoinOperation);
            oExtrudeDef.SetDistanceExtent(3, Inventor.PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
        }

        private void Feature4(int x, int y)
        {
            try
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
                oPartDoc.UnitsOfMeasure.LengthUnits = Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;
                //Inventor.Point2d point1 = _invApp.TransientGeometry.CreatePoint2d(6, 13.8);
                //Inventor.Point2d point2 = _invApp.TransientGeometry.CreatePoint2d(-6, 13.8);
                PlanarSketch oSketch = default(PlanarSketch);
                oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);
                SketchArc arc1 = oSketch.SketchArcs.AddByThreePoints(_invApp.TransientGeometry.CreatePoint2d(5.879, 13.8), _invApp.TransientGeometry.CreatePoint2d(0, 15), _invApp.TransientGeometry.CreatePoint2d(-5.879, 13.8));
                SketchArc arc2 = oSketch.SketchArcs.AddByThreePoints(arc1.StartSketchPoint, _invApp.TransientGeometry.CreatePoint2d(0, 9), arc1.EndSketchPoint);
                //oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 0), 15);
                //oSketch.SketchCircles.AddByCenterRadius(_invApp.TransientGeometry.CreatePoint2d(0, 15), 6);

                Profile oProfile = default(Profile);
                oProfile = oSketch.Profiles.AddForSolid();
                //oProfile = oSketch.Profiles.AddForSurface(arc2);

                ExtrudeDefinition oExtrudeDef = default(ExtrudeDefinition);
                oExtrudeDef = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, Inventor.PartFeatureOperationEnum.kCutOperation);
                oExtrudeDef.SetDistanceExtent(3, Inventor.PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

                ExtrudeFeature oExtrude = default(ExtrudeFeature);
                oExtrude = oCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef);
            }
            catch
            {
                try
                {
                    Feature3(x, y);
                }
                catch
                {
                    Feature2(x, y);
                }
            }
            
        }

        private void Feature5(int x, int y)
        {
            try
            {
                PartDocument oPartDoc = (PartDocument)_invApp.ActiveDocument;
                oPartDoc.UnitsOfMeasure.LengthUnits = Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
                PartComponentDefinition oCompDef = default(PartComponentDefinition);
                oCompDef = oPartDoc.ComponentDefinition;
                Inventor.ObjectCollection objCol = default(Inventor.ObjectCollection);
                objCol = _invApp.TransientObjects.CreateObjectCollection();
                objCol.Add(oCompDef.Features[4]);
                oCompDef.Features.CircularPatternFeatures.Add(ParentFeatures: objCol, AxisEntity: oCompDef.WorkAxes[3], NaturalAxisDirection: true, Count: 6, Angle: 360 * 0.0174532925, FitWithinAngle: true, ComputeType: Inventor.PatternComputeTypeEnum.kIdenticalCompute);

                serial1.Close();
                pictureBox2.Enabled = false;
                label2.Enabled = false;
                pictureBox1.Enabled = true;
                label1.Enabled = true;

                MessageBox.Show(this,
                  "Good job! Your design is finished",
                  "Finish",
                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            catch
            {
                try
                {
                    Feature4(x, y);
                }
                catch
                {
                    try
                    {
                        Feature3(x, y);
                    }
                    catch
                    {
                        Feature2(x, y);
                    }
                }

            }
           
        }

        private void SaveCSV(int x, int y)
        {
            StringBuilder csvcontent = new StringBuilder();
            string csvpath = "D:\\CSVFile\\ImprosData\\XYData_1.csv";

            //var newline = String.Format("{0},{1},{2},{3},{4},{5},{6}", DateTime.Now.ToString("HH:mm:ss.fff tt"), "X", x.ToString(), " ", "y", y.ToString(), System.Environment.NewLine);
            var newline = String.Format("{0},{1}", DateTime.Now.ToString("ss.fff tt"), System.Environment.NewLine);
            csvcontent.Append(newline);
            System.IO.File.AppendAllText(csvpath, csvcontent.ToString());
        }

        private void SaveCSV_SerialRead()
        {
            StringBuilder csvcontent = new StringBuilder();
            string csvpath = "D:\\CSVFile\\ImprosData\\SerialRead_Calc.csv";

            var newline = String.Format("{0},{1}", DateTime.Now.ToString("HH:mm:ss tt"), System.Environment.NewLine);

            System.IO.File.AppendAllText(csvpath, csvcontent.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //ExtrudeDoub(19, 30);
                //CircleSketch(6,8);
                
                movrot.GetCurrentCamera();
                //movrot.ChangeView(int.Parse(textBox2.Text), int.Parse(textBox3.Text), int.Parse(textBox4.Text), apply: true);
                //movrot.TranslateView(int.Parse(textBox2.Text), int.Parse(textBox3.Text));
            }
            catch (System.FormatException)
            {
                return;
            }
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            _invApp.UserInterfaceManager.RibbonDockingState = RibbonDockingStateEnum.kDockToTop;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /*stopwatch.Start();
            Feature1();
            Feature2();
            Feature3();
            Feature4();
            Feature5();
            Thread.Sleep(500);
            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);*/
            
        }
    }
}
