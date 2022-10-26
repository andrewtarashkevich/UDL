using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Catalogs;
using Tekla.Structures.Datatype;
using Tekla.Structures.Dialog;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Drawing;
using Point = Tekla.Structures.Geometry3d.Point;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;
using MoreLinq;
using System.Globalization;

namespace UDL
{
    public partial class UDL : Form
    {
        public UDL()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(1450, 150);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Model model = new Model();

            ProjectInfo method = model.GetProjectInfo();
            if (method.Info1 == "ASD")
            {
                textBox3.Text = Convert.ToString("ASD");
            }
            else if (method.Info1 == "LRFD")
            {
                textBox3.Text = Convert.ToString("LRFD");
            }
            else
            {
                textBox3.Text = Convert.ToString("Info1->ASD/LRFD");
            }

            if (model.GetConnectionStatus() == false) return;

            ProjectInfo perc = model.GetProjectInfo();

            if (perc.Info2 == "" || perc.Info2 == "PROJ-INFO_2")
            {
                textBox4.Text = "Info2->Enter number";
            }
            else
            {         
                double percent = Convert.ToDouble(perc.Info2) / 100;
                textBox4.Text = perc.Info2;

                Picker BeamPicker = new Picker();
                Point Point1 = BeamPicker.PickPoint("Pick 1st point");
                Point Point2 = BeamPicker.PickPoint("Pick 2nd point");

                double lg_real = Math.Sqrt((Point2.X - Point1.X) * (Point2.X - Point1.X) + (Point2.Y - Point1.Y) * (Point2.Y - Point1.Y)) / 25.4;
                double lg_res = 12 * Math.Truncate((Math.Round((Math.Sqrt((Point2.X - Point1.X) * (Point2.X - Point1.X) + (Point2.Y - Point1.Y) * (Point2.Y - Point1.Y)) / 25.4), 3)) / 12);
                textBox2.Text = Convert.ToString(Math.Round((lg_real), 3));

                Picker p = new Picker();
                Tekla.Structures.Model.ModelObject mo = p.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, "Pick beam");

                string fileName = "C:\\Dropbox\\technology\\Utilits\\UDL\\aisc.database.xlsx"; //aisc shape database folder
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);

                if (mo is Beam b)
                {
                    string pr = "PROFILE";
                    string pr_res = "";
                    b.GetReportProperty(pr, ref pr_res);

                    foreach (var column in worksheet.RowsUsed().Where(c => c.FirstCell().GetString() == pr_res))
                    {

                        double Zx_res = Convert.ToDouble(column.Cell(18).Value);
                        double d_res = Convert.ToDouble(column.Cell(3).Value);
                        double bf_res = Convert.ToDouble(column.Cell(5).Value);
                        double tf_res = Convert.ToDouble(column.Cell(6).Value);
                        double tw_res = Convert.ToDouble(column.Cell(4).Value);
                        double Sx_res = Convert.ToDouble(column.Cell(16).Value);
                        double h_res = Convert.ToDouble(column.Cell(10).Value);
                        textBox1.Text = Convert.ToString(pr_res);

                        double eq1 = bf_res / (2 * tf_res);
                        double eq2 = h_res / tw_res;
                        double fi1 = 0.9; // LRFD
                        double fi2 = 1.0; // LRFD
                        double omega1 = 1.5; // ASD
                        double omega2 = 1.67; // ASD


                        //ASD
                        double McA = 50 * Zx_res / omega2; // for compact flanges
                        double MncA = ((50 * Zx_res - (50 * Zx_res - 0.7 * 50 * Sx_res) * ((eq1 - 9.1516) / (24.0832 - 9.1516)))) / omega2; // for noncompact flanges
                        double VlA = 0.6 * 50 * d_res * tw_res / omega1; // for h/tw <=53.946
                        double VmA = 0.6 * 50 * d_res * tw_res / omega2; // for h/tw >53.946
                        double LlimA;
                        double Llim1A = 4 * McA / VlA;
                        double Llim2A = 4 * McA / VmA;
                        double Llim3A = 4 * MncA / VlA;
                        double q_resA;
                        double q_mmaxA;
                        double qcA = McA * 8 / lg_res;
                        double qncA = MncA * 8 / lg_res;
                        double qlA = 2 * VlA;
                        double qmA = 2 * VmA;

                        if (eq1 <= 9.1516 && eq2 <= 53.946)
                        {
                            LlimA = Llim1A;
                        } // get Llim 
                        else
                        {
                            if (eq1 <= 9.1516 && eq2 > 53.946)
                            {
                                LlimA = Llim2A;
                            }
                            else
                            {
                                LlimA = Llim3A;
                            }
                        }

                        if (eq1 <= 9.1516 && eq2 <= 53.946)
                        {
                            if (lg_res > (LlimA))
                            {
                                q_resA = qcA;
                            }
                            else
                            {
                                q_resA = qlA;
                            }
                        } // get uniform dead load q_res
                        else
                        {
                            if (eq1 <= 9.1516 && eq2 > 53.946)
                            {
                                if (lg_res > (LlimA))
                                {
                                    q_resA = qcA;
                                }
                                else
                                {
                                    q_resA = qmA;
                                }
                            }
                            else
                            {
                                if (lg_res > (LlimA))
                                {
                                    q_resA = qncA;
                                }
                                else
                                {
                                    q_resA = qlA;
                                }
                            }
                        }

                        if (eq2 <= 53.946)
                        {
                            q_mmaxA = qlA;
                        } // for Max. Moment
                        else
                        {
                            q_mmaxA = qmA;
                        }



                        //LRFD
                        double McL = fi1 * 50 * Zx_res; // for compact flanges
                        double MncL = fi1 * (50 * Zx_res - (50 * Zx_res - 0.7 * 50 * Sx_res) * ((eq1 - 9.1516) / (24.0832 - 9.1516))); // for noncompact flanges
                        double VlL = fi2 * 0.6 * 50 * d_res * tw_res; // for h/tw <=53.946
                        double VmL = fi1 * 0.6 * 50 * d_res * tw_res; // for h/tw >53.946
                        double LlimL;
                        double Llim1L = 4 * McL / VlL;
                        double Llim2L = 4 * McL / VmL;
                        double Llim3L = 4 * MncL / VlL;
                        double q_resL;
                        double q_mmaxL;
                        double qcL = McL * 8 / lg_res;
                        double qncL = MncL * 8 / lg_res;
                        double qlL = 2 * VlL;
                        double qmL = 2 * VmL;

                        if (eq1 <= 9.1516 && eq2 <= 53.946)
                        {
                            LlimL = Llim1L;
                        } // get Llim 
                        else
                        {
                            if (eq1 <= 9.1516 && eq2 > 53.946)
                            {
                                LlimL = Llim2L;
                            }
                            else
                            {
                                LlimL = Llim3L;
                            }
                        }

                        if (eq1 <= 9.1516 && eq2 <= 53.946)
                        {
                            if (lg_res > (LlimL))
                            {
                                q_resL = qcL;
                            }
                            else
                            {
                                q_resL = qlL;
                            }
                        } // get uniform dead load q_res
                        else
                        {
                            if (eq1 <= 9.1516 && eq2 > 53.946)
                            {
                                if (lg_res > (LlimL))
                                {
                                    q_resL = qcL;
                                }
                                else
                                {
                                    q_resL = qmL;
                                }
                            }
                            else
                            {
                                if (lg_res > (LlimL))
                                {
                                    q_resL = qncL;
                                }
                                else
                                {
                                    q_resL = qlL;
                                }
                            }
                        }

                        if (eq2 <= 53.946)
                        {
                            q_mmaxL = qlL;
                        } // for Max. Moment
                        else
                        {
                            q_mmaxL = qmL;
                        }

                        if (method.Info1 == "ASD")
                        {
                            textBox10.Text = Convert.ToString(Math.Round(LlimA / (12), 2)); // Limiting Span for Shear Strength, ft
                            textBox11.Text = Convert.ToString(Math.Round(q_mmaxA * LlimA / (8 * 12), 2)); // Max. Moment,  kip-ft
                            textBox12.Text = Convert.ToString(Math.Round(q_resA * percent, 2)); // Reaction,  kips
                            b.SetUserProperty("shear1", q_resA * 4448 * percent);
                            b.SetUserProperty("shear2", q_resA * 4448 * percent);
                            b.SetUserProperty("moment1", q_mmaxA * 1355.86 * LlimA / (8 * 12));
                            b.SetUserProperty("moment2", q_mmaxA * 1355.86 * LlimA / (8 * 12));
                        }
                        else if (method.Info1 == "LRFD")
                        {
                            textBox10.Text = Convert.ToString(Math.Round(LlimL / (12), 2)); // Limiting Span for Shear Strength, ft
                            textBox11.Text = Convert.ToString(Math.Round(q_mmaxL * LlimL / (8 * 12), 2)); // Max. Moment,  kip-ft
                            textBox12.Text = Convert.ToString(Math.Round(q_resL * percent, 2)); // Reaction,  kips
                            b.SetUserProperty("shear1", q_resL * percent / 2);
                            b.SetUserProperty("shear2", q_resL * percent / 2);
                            b.SetUserProperty("moment1", q_mmaxL * 1355.86 * LlimA / (8 * 12));
                            b.SetUserProperty("moment2", q_mmaxL * 1355.86 * LlimA / (8 * 12));
                        }
                        else
                        {
                            textBox3.Text = Convert.ToString("Info1->ASD/LRFD");
                        }


                    }

                }
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
            
               
        private void button3_Click(object sender, EventArgs e)
        {
            Model model = new Model();

            ProjectInfo method = model.GetProjectInfo();
            if (method.Info1 == "ASD")
            {
                textBox3.Text = Convert.ToString("ASD");
            }
            else if (method.Info1 == "LRFD")
            {
                textBox3.Text = Convert.ToString("LRFD");
            }
            else
            {
                textBox3.Text = Convert.ToString("Info1->ASD/LRFD");
                return;
            }

            if (model.GetConnectionStatus() == false) return;

            ProjectInfo perc = model.GetProjectInfo();

            if (perc.Info2 == "" || perc.Info2 == "PROJ-INFO_2")
            {
                textBox4.Text = "Info2->Enter number";
            }
            else
            {
                double percent = Convert.ToDouble(perc.Info2) / 100;
                textBox4.Text = perc.Info2;

                Picker p = new Picker();

                string fileName = "C:\\Dropbox\\technology\\Utilits\\UDL\\aisc.database.xlsx"; //aisc shape database folder
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);

                ModelObjectEnumerator mo = model.GetModelObjectSelector().GetObjectsByFilterName("FF_UDL");

                while (mo.MoveNext())
                {
                    if (mo.Current is Beam b)
                    {
                        Point point = new Point();
                        Point point2 = new Point();
                        point2 = b.StartPoint;
                        point = b.EndPoint;

                        double lg_res = 12 * Math.Truncate((Math.Round((Math.Sqrt(Math.Pow(point2.X - point.X, 2) + Math.Pow(point2.Y - point.Y, 2)) / 25.4), 3)) / 12);

                        string pr = "PROFILE";
                        string pr_res = "";
                        b.GetReportProperty(pr, ref pr_res);

                        foreach (var column in worksheet.RowsUsed().Where(c => c.FirstCell().GetString() == pr_res))
                        {

                            double Zx_res = Convert.ToDouble(column.Cell(18).Value);
                            double d_res = Convert.ToDouble(column.Cell(3).Value);
                            double bf_res = Convert.ToDouble(column.Cell(5).Value);
                            double tf_res = Convert.ToDouble(column.Cell(6).Value);
                            double tw_res = Convert.ToDouble(column.Cell(4).Value);
                            double Sx_res = Convert.ToDouble(column.Cell(16).Value);
                            double h_res = Convert.ToDouble(column.Cell(10).Value);


                            double eq1 = bf_res / (2 * tf_res);
                            double eq2 = h_res / tw_res;
                            double fi1 = 0.9; // LRFD
                            double fi2 = 1.0; // LRFD
                            double omega1 = 1.5; // ASD
                            double omega2 = 1.67; // ASD


                            //ASD
                            double McA = 50 * Zx_res / omega2; // for compact flanges
                            double MncA = ((50 * Zx_res - (50 * Zx_res - 0.7 * 50 * Sx_res) * ((eq1 - 9.1516) / (24.0832 - 9.1516)))) / omega2; // for noncompact flanges
                            double VlA = 0.6 * 50 * d_res * tw_res / omega1; // for h/tw <=53.946
                            double VmA = 0.6 * 50 * d_res * tw_res / omega2; // for h/tw >53.946
                            double LlimA;
                            double Llim1A = 4 * McA / VlA;
                            double Llim2A = 4 * McA / VmA;
                            double Llim3A = 4 * MncA / VlA;
                            double q_resA;
                            double q_mmaxA;
                            double qcA = McA * 8 / lg_res;
                            double qncA = MncA * 8 / lg_res;
                            double qlA = 2 * VlA;
                            double qmA = 2 * VmA;

                            if (eq1 <= 9.1516 && eq2 <= 53.946)
                            {
                                LlimA = Llim1A;
                            } // get Llim 
                            else
                            {
                                if (eq1 <= 9.1516 && eq2 > 53.946)
                                {
                                    LlimA = Llim2A;
                                }
                                else
                                {
                                    LlimA = Llim3A;
                                }
                            }

                            if (eq1 <= 9.1516 && eq2 <= 53.946)
                            {
                                if (lg_res > (LlimA))
                                {
                                    q_resA = qcA;
                                }
                                else
                                {
                                    q_resA = qlA;
                                }
                            } // get uniform dead load q_res
                            else
                            {
                                if (eq1 <= 9.1516 && eq2 > 53.946)
                                {
                                    if (lg_res > (LlimA))
                                    {
                                        q_resA = qcA;
                                    }
                                    else
                                    {
                                        q_resA = qmA;
                                    }
                                }
                                else
                                {
                                    if (lg_res > (LlimA))
                                    {
                                        q_resA = qncA;
                                    }
                                    else
                                    {
                                        q_resA = qlA;
                                    }
                                }
                            }

                            if (eq2 <= 53.946)
                            {
                                q_mmaxA = qlA;
                            } // for Max. Moment
                            else
                            {
                                q_mmaxA = qmA;
                            }



                            //LRFD
                            double McL = fi1 * 50 * Zx_res; // for compact flanges
                            double MncL = fi1 * (50 * Zx_res - (50 * Zx_res - 0.7 * 50 * Sx_res) * ((eq1 - 9.1516) / (24.0832 - 9.1516))); // for noncompact flanges
                            double VlL = fi2 * 0.6 * 50 * d_res * tw_res; // for h/tw <=53.946
                            double VmL = fi1 * 0.6 * 50 * d_res * tw_res; // for h/tw >53.946
                            double LlimL;
                            double Llim1L = 4 * McL / VlL;
                            double Llim2L = 4 * McL / VmL;
                            double Llim3L = 4 * MncL / VlL;
                            double q_resL;
                            double q_mmaxL;
                            double qcL = McL * 8 / lg_res;
                            double qncL = MncL * 8 / lg_res;
                            double qlL = 2 * VlL;
                            double qmL = 2 * VmL;

                            if (eq1 <= 9.1516 && eq2 <= 53.946)
                            {
                                LlimL = Llim1L;
                            } // get Llim 
                            else
                            {
                                if (eq1 <= 9.1516 && eq2 > 53.946)
                                {
                                    LlimL = Llim2L;
                                }
                                else
                                {
                                    LlimL = Llim3L;
                                }
                            }

                            if (eq1 <= 9.1516 && eq2 <= 53.946)
                            {
                                if (lg_res > (LlimL))
                                {
                                    q_resL = qcL;
                                }
                                else
                                {
                                    q_resL = qlL;
                                }
                            } // get uniform dead load q_res
                            else
                            {
                                if (eq1 <= 9.1516 && eq2 > 53.946)
                                {
                                    if (lg_res > (LlimL))
                                    {
                                        q_resL = qcL;
                                    }
                                    else
                                    {
                                        q_resL = qmL;
                                    }
                                }
                                else
                                {
                                    if (lg_res > (LlimL))
                                    {
                                        q_resL = qncL;
                                    }
                                    else
                                    {
                                        q_resL = qlL;
                                    }
                                }
                            }

                            if (eq2 <= 53.946)
                            {
                                q_mmaxL = qlL;
                            } // for Max. Moment
                            else
                            {
                                q_mmaxL = qmL;
                            }

                            if (method.Info1 == "ASD")
                            {
                                b.SetUserProperty("shear1", q_resA * 4448 * percent);
                                b.SetUserProperty("shear2", q_resA * 4448 * percent);
                                b.SetUserProperty("moment1", q_mmaxA * 1355.86 * LlimA / (8 * 12));
                                b.SetUserProperty("moment2", q_mmaxA * 1355.86 * LlimA / (8 * 12));
                            }
                            else if (method.Info1 == "LRFD")
                            {
                                b.SetUserProperty("shear1", q_resL * 4448 * percent);
                                b.SetUserProperty("shear2", q_resL * 4448 * percent);
                                b.SetUserProperty("moment1", q_mmaxL * 1355.86 * LlimA / (8 * 12));
                                b.SetUserProperty("moment2", q_mmaxL * 1355.86 * LlimA / (8 * 12));
                            }
                            else
                            {
                                textBox3.Text = Convert.ToString("Info1->ASD/LRFD");
                            }


                        }

                    }
                }

                textBox12.Text = Convert.ToString("Complete");
            }

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Tekla.Structures.Model.Model MyModel = new Tekla.Structures.Model.Model();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            if (MyModel.GetConnectionStatus())
            {
                if (MyDrawingHandler.GetConnectionStatus())
                {
                    DrawingHandler.SetMessageExecutionStatus(DrawingHandler.MessageExecutionModeEnum.INSTANT);
                    Drawing CurrentDrawing = MyDrawingHandler.GetActiveDrawing();                   

                    DrawingObjectEnumerator mo = CurrentDrawing.GetSheet().GetAllObjects(typeof(Tekla.Structures.Drawing.Part));
                    
                    while (mo.MoveNext())
                    {
                        if (mo.Current is Tekla.Structures.Drawing.Part)
                        {
                            Tekla.Structures.Drawing.Part myPart = (Tekla.Structures.Drawing.Part)mo.Current;
                            Tekla.Structures.Identifier Identifier = myPart.ModelIdentifier;
                            Tekla.Structures.Model.ModelObject ModelSideObject = new Model().SelectModelObject(Identifier);

                            PointList PointList = new PointList();
                            Beam myBeam = new Beam();
                            if (ModelSideObject.GetType().Equals(typeof(Beam)))
                            {
                                myBeam.Identifier = Identifier;
                                myBeam.Select();
                                string pr = "PROFILE";
                                string pr_w = "W";
                                string pr_res = "";
                                myBeam.GetReportProperty(pr, ref pr_res);
                                string nm = "NAME";
                                string nm_b = "B";
                                string nm_e = "E";
                                string nm_res = "";
                                myBeam.GetReportProperty(nm, ref nm_res);
                                int indexOfpr = pr_res.IndexOf(pr_w);
                                if (indexOfpr == 0)
                                {
                                    int indexOfnm0 = nm_res.IndexOf(nm_b);
                                    int indexOfnm1 = nm_res.IndexOf(nm_e);
                                    if (indexOfnm0 == 0 & indexOfnm1 == 1)
                                    {
                                        PointList.Add(myBeam.StartPoint);
                                        PointList.Add(myBeam.EndPoint);

                                        Tekla.Structures.Drawing.ModelObject modelObject = (Tekla.Structures.Drawing.ModelObject)mo.Current;

                                        //Point PartMinPoint = null, PartMaxPoint = null, PartEnd1Point = null, PartEnd2Point = null, PartCenterPoint = null;
                                        //GetPartPoints(MyModel, modelObject.GetView(), modelObject, out PartMinPoint, out PartMaxPoint, out PartEnd1Point, out PartEnd2Point, out PartCenterPoint);

                                        Mark Mark1 = new Mark(modelObject);
                                        Mark1.Attributes.Content.Clear();
                                        Mark1.Attributes.Content.Add(new TextElement("B-"));
                                        Mark1.Attributes.Content.Add(new NewLineElement());
                                        Mark1.Attributes.Content.Add(new TextElement("V="));
                                        Mark1.Attributes.Content.Add(new UserDefinedElement("shear1"));
                                        Mark1.Attributes.Content.Add(new TextElement("kips"));
                                        Mark1.Placing = new LeaderLinePlacing(myBeam.StartPoint);
                                        Mark1.Attributes.PlacingAttributes.IsFixed = false;
                                        //Mark1.InsertionPoint = PartCenterPoint;
                                        Mark1.Insert();


                                        Mark Mark2 = new Mark(modelObject);
                                        Mark2.Attributes.Content.Clear();
                                        Mark2.Attributes.Content.Add(new TextElement("B-"));
                                        Mark2.Attributes.Content.Add(new NewLineElement());
                                        Mark2.Attributes.Content.Add(new TextElement("V="));
                                        Mark2.Attributes.Content.Add(new UserDefinedElement("shear2"));
                                        Mark2.Attributes.Content.Add(new TextElement("kips"));
                                        Mark2.Placing = new LeaderLinePlacing(myBeam.EndPoint);
                                        Mark2.Attributes.PlacingAttributes.IsFixed = false;
                                        //Mark2.InsertionPoint = PartCenterPoint;
                                        Mark2.Insert();
                                    }
                                }

                            }
                        }

                    }
                }
            }
        }

        private void GetPartPoints(Tekla.Structures.Model.Model MyModel, ViewBase PartView, Tekla.Structures.Drawing.ModelObject modelObject, out Point PartMinPoint, out Point PartMaxPoint, out Point PartEnd1Point, out Point PartEnd2Point, out Point PartCenterPoint)
        {
            Tekla.Structures.Model.ModelObject modelPart = GetModelObjectFromDrawingModelObject(MyModel, modelObject);
            GetModelObjectStartAndEndPoint(modelPart, (Tekla.Structures.Drawing.View)PartView, out PartMinPoint, out PartMaxPoint);
            PartEnd1Point = GetInsertion1Point(PartMinPoint, PartMaxPoint);
            PartEnd2Point = GetInsertion2Point(PartMinPoint, PartMaxPoint);
            PartCenterPoint = GetInsertion3Point(PartMinPoint, PartMaxPoint);
        }

        private Tekla.Structures.Model.ModelObject GetModelObjectFromDrawingModelObject(Tekla.Structures.Model.Model MyModel, Tekla.Structures.Drawing.ModelObject PartOfMark)
        {
            Tekla.Structures.Model.ModelObject modelObject = MyModel.SelectModelObject(PartOfMark.ModelIdentifier);

            Tekla.Structures.Model.Part modelPart = (Tekla.Structures.Model.Part)modelObject;

            return modelPart;
        }

        private void GetModelObjectStartAndEndPoint(Tekla.Structures.Model.ModelObject modelObject, Tekla.Structures.Drawing.View PartView, out Point PartStartPoint, out Point PartEndPoint)
        {
            Tekla.Structures.Model.Part modelPart = (Tekla.Structures.Model.Part)modelObject;

            PartStartPoint = modelPart.GetSolid().MinimumPoint;
            PartEndPoint = modelPart.GetSolid().MaximumPoint;

            Matrix convMatrix = MatrixFactory.ToCoordinateSystem(PartView.DisplayCoordinateSystem);
            PartStartPoint = convMatrix.Transform(PartStartPoint);
            PartEndPoint = convMatrix.Transform(PartEndPoint);
        }

        private Point GetInsertion1Point(Point PartStartPoint, Point PartEndPoint)
        {
            Point MinPoint = PartStartPoint;
            Point MaxPoint = PartEndPoint;
            Point InsertionPoint = new Point(MinPoint.X, MaxPoint.Y);
            return InsertionPoint;
        }

        private Point GetInsertion2Point(Point PartStartPoint, Point PartEndPoint)
        {
            Point MinPoint = PartStartPoint;
            Point MaxPoint = PartEndPoint;
            Point InsertionPoint = new Point(MaxPoint.X, MinPoint.Y);
            return InsertionPoint;
        }

        private Point GetInsertion3Point(Point PartStartPoint, Point PartEndPoint)
        {
            Point MinPoint = PartStartPoint;
            Point MaxPoint = PartEndPoint;
            Point InsertionPoint = new Point((MaxPoint.X + MinPoint.X) * 0.5, (MaxPoint.Y + MinPoint.Y) * 0.5);
            return InsertionPoint;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

}

