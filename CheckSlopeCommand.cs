using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace CheckSlopePipe
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CheckSlopeCommand : IExternalCommand
    {
        // Tham số cấu hình độ dốc
        private const double Diameter_1 = 100.0; // Đường kính 1 (mm)
        private const double Slope_1 = 0.02;     // Độ dốc 1 (2%)
        private const double Diameter_2 = 50.0;  // Đường kính 2 (mm)
        private const double Slope_2 = 0.01;     // Độ dốc 2 (1%)

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Hiển thị hộp thoại nhập thông số
                SlopeParametersForm form = new SlopeParametersForm();
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Lấy tất cả đường ống trong dự án
                    List<Pipe> allPipes = GetPipes(doc);
                    
                    // Lọc các đường ống không đạt yêu cầu
                    List<Pipe> nonCompliantPipes = CheckPipeSlopes(allPipes, form.Diameter1, form.Slope1, form.Diameter2, form.Slope2);
                    
                    // Tạo bảng dự toán
                    GenerateQuantityTable(doc, nonCompliantPipes, form.Diameter1, form.Slope1, form.Diameter2, form.Slope2);
                    
                    // Hiển thị kết quả
                    TaskDialog.Show("Kết quả kiểm tra", 
                        $"Đã kiểm tra {allPipes.Count} đường ống.\n" +
                        $"Tìm thấy {nonCompliantPipes.Count} đường ống không đạt yêu cầu.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Lỗi", $"Đã xảy ra lỗi: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Lấy tất cả đường ống từ tài liệu Revit
        /// </summary>
        private List<Pipe> GetPipes(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Pipe));
            
            List<Pipe> pipes = new List<Pipe>();
            foreach (Element element in collector)
            {
                Pipe pipe = element as Pipe;
                if (pipe != null)
                {
                    pipes.Add(pipe);
                }
            }
            
            return pipes;
        }

        /// <summary>
        /// Kiểm tra độ dốc của đường ống theo điều kiện
        /// </summary>
        private List<Pipe> CheckPipeSlopes(List<Pipe> pipes, double diameter1, double slope1, double diameter2, double slope2)
        {
            List<Pipe> nonCompliantPipes = new List<Pipe>();

            foreach (Pipe pipe in pipes)
            {
                // Bỏ qua đường ống thẳng đứng
                if (IsVerticalPipe(pipe))
                    continue;

                // Lấy đường kính đường ống (chuyển đổi từ feet sang mm)
                double pipeDiameter = GetPipeDiameter(pipe) * 304.8; // Chuyển feet sang mm

                // Kiểm tra điều kiện độ dốc
                double pipeSlope = GetPipeSlope(pipe);
                
                bool isCompliant = false;
                
                if (pipeDiameter >= diameter1)
                {
                    // Kiểm tra độ dốc cho đường kính lớn
                    isCompliant = Math.Abs(pipeSlope - slope1) < 0.001; // Dung sai nhỏ
                }
                else if (pipeDiameter <= diameter2)
                {
                    // Kiểm tra độ dốc cho đường kính nhỏ
                    isCompliant = Math.Abs(pipeSlope - slope2) < 0.001; // Dung sai nhỏ
                }
                else
                {
                    // Đường kính nằm giữa hai giá trị - không cần kiểm tra
                    isCompliant = true;
                }

                if (!isCompliant)
                {
                    nonCompliantPipes.Add(pipe);
                }
            }

            return nonCompliantPipes;
        }

        /// <summary>
        /// Kiểm tra xem đường ống có thẳng đứng không
        /// </summary>
        private bool IsVerticalPipe(Pipe pipe)
        {
            LocationCurve location = pipe.Location as LocationCurve;
            if (location == null) return false;

            Curve curve = location.Curve;
            if (curve == null) return false;

            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);

            // Kiểm tra nếu đường ống gần như thẳng đứng (chênh lệch Z lớn, X,Y nhỏ)
            double deltaZ = Math.Abs(endPoint.Z - startPoint.Z);
            double deltaXY = Math.Sqrt(Math.Pow(endPoint.X - startPoint.X, 2) + Math.Pow(endPoint.Y - startPoint.Y, 2));

            // Nếu deltaZ lớn hơn nhiều so với deltaXY, coi là đường ống thẳng đứng
            return deltaZ > deltaXY * 10; // Ngưỡng có thể điều chỉnh
        }

        /// <summary>
        /// Lấy đường kính đường ống (tính bằng feet)
        /// </summary>
        private double GetPipeDiameter(Pipe pipe)
        {
            Parameter diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (diameterParam != null && diameterParam.HasValue)
            {
                return diameterParam.AsDouble();
            }
            return 0.0;
        }

        /// <summary>
        /// Lấy độ dốc của đường ống
        /// </summary>
        private double GetPipeSlope(Pipe pipe)
        {
            Parameter slopeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
            if (slopeParam != null && slopeParam.HasValue)
            {
                return slopeParam.AsDouble();
            }
            return 0.0;
        }

        /// <summary>
        /// Tạo bảng dự toán cho các đường ống không đạt yêu cầu
        /// </summary>
        private void GenerateQuantityTable(Document doc, List<Pipe> nonCompliantPipes, double diameter1, double slope1, double diameter2, double slope2)
        {
            // Tạo transaction để thêm dữ liệu vào dự án
            using (Transaction trans = new Transaction(doc, "Tạo bảng dự toán"))
            {
                trans.Start();

                try
                {
                    // Tạo schedule cho đường ống
                    ViewSchedule schedule = CreatePipeSchedule(doc);
                    
                    // Thêm các đường ống không đạt yêu cầu vào schedule
                    AddPipesToSchedule(doc, schedule, nonCompliantPipes);
                    
                    trans.Commit();
                    
                    TaskDialog.Show("Thông báo", "Đã tạo bảng dự toán cho các đường ống không đạt yêu cầu.");
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new Exception($"Lỗi khi tạo bảng dự toán: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Tạo schedule mới cho đường ống
        /// </summary>
        private ViewSchedule CreatePipeSchedule(Document doc)
        {
            // Lấy category đường ống
            Category pipeCategory = Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves);
            
            // Tạo schedule
            ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, pipeCategory.Id);
            schedule.Name = "Bảng dự toán đường ống không đạt độ dốc";

            // Thêm các trường cần thiết vào schedule
            AddScheduleFields(doc, schedule);

            return schedule;
        }

        /// <summary>
        /// Thêm các trường vào schedule
        /// </summary>
        private void AddScheduleFields(Document doc, ViewSchedule schedule)
        {
            // Lấy định nghĩa schedule
            ScheduleDefinition definition = schedule.Definition;

            // Thêm các trường thông tin đường ống
            AddScheduleField(definition, BuiltInParameter.ALL_MODEL_TYPE_NAME); // Loại đường ống
            AddScheduleField(definition, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM); // Đường kính
            AddScheduleField(definition, BuiltInParameter.RBS_PIPE_SLOPE); // Độ dốc
            AddScheduleField(definition, BuiltInParameter.RBS_START_LEVEL_PARAM); // Cao độ đầu
            AddScheduleField(definition, BuiltInParameter.RBS_END_LEVEL_PARAM); // Cao độ cuối
            AddScheduleField(definition, BuiltInParameter.CURVE_ELEM_LENGTH); // Chiều dài
        }

        /// <summary>
        /// Thêm một trường vào schedule
        /// </summary>
        private void AddScheduleField(ScheduleDefinition definition, BuiltInParameter parameter)
        {
            ScheduleFieldId fieldId = definition.AddField(parameter);
            // Có thể tùy chỉnh thêm định dạng trường ở đây
        }

        /// <summary>
        /// Thêm đường ống vào schedule
        /// </summary>
        private void AddPipesToSchedule(Document doc, ViewSchedule schedule, List<Pipe> pipes)
        {
            // Trong Revit, schedule tự động hiển thị các element dựa trên bộ lọc
            // Chúng ta có thể tạo bộ lọc để chỉ hiển thị các đường ống không đạt yêu cầu
            ScheduleDefinition definition = schedule.Definition;

            // Xóa các bộ lọc hiện có
            IList<ScheduleFilter> filters = definition.GetFilters();
            foreach (ScheduleFilter filter in filters)
            {
                definition.RemoveFilter(filter);
            }

            // Tạo bộ lọc mới (nếu cần)
            // Trong trường hợp này, schedule sẽ hiển thị tất cả đường ống được chọn
        }
    }

    /// <summary>
    /// Form nhập thông số độ dốc
    /// </summary>
    public partial class SlopeParametersForm : System.Windows.Forms.Form
    {
        public double Diameter1 { get; private set; }
        public double Slope1 { get; private set; }
        public double Diameter2 { get; private set; }
        public double Slope2 { get; private set; }

        private System.Windows.Forms.TextBox txtDiameter1;
        private System.Windows.Forms.TextBox txtSlope1;
        private System.Windows.Forms.TextBox txtDiameter2;
        private System.Windows.Forms.TextBox txtSlope2;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;

        public SlopeParametersForm()
        {
            InitializeComponent();
            LoadDefaultValues();
        }

        private void InitializeComponent()
        {
            this.txtDiameter1 = new System.Windows.Forms.TextBox();
            this.txtSlope1 = new System.Windows.Forms.TextBox();
            this.txtDiameter2 = new System.Windows.Forms.TextBox();
            this.txtSlope2 = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // label1
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.Text = "Đường kính 1 (mm):";

            // txtDiameter1
            this.txtDiameter1.Location = new System.Drawing.Point(150, 17);
            this.txtDiameter1.Name = "txtDiameter1";
            this.txtDiameter1.Size = new System.Drawing.Size(100, 20);
            
            // label2
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 13);
            this.label2.Text = "Độ dốc 1 (ví dụ: 0.02):";

            // txtSlope1
            this.txtSlope1.Location = new System.Drawing.Point(150, 47);
            this.txtSlope1.Name = "txtSlope1";
            this.txtSlope1.Size = new System.Drawing.Size(100, 20);
            
            // label3
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 13);
            this.label3.Text = "Đường kính 2 (mm):";

            // txtDiameter2
            this.txtDiameter2.Location = new System.Drawing.Point(150, 77);
            this.txtDiameter2.Name = "txtDiameter2";
            this.txtDiameter2.Size = new System.Drawing.Size(100, 20);
            
            // label4
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 110);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 13);
            this.label4.Text = "Độ dốc 2 (ví dụ: 0.01):";

            // txtSlope2
            this.txtSlope2.Location = new System.Drawing.Point(150, 107);
            this.txtSlope2.Name = "txtSlope2";
            this.txtSlope2.Size = new System.Drawing.Size(100, 20);
            
            // btnOK
            this.btnOK.Location = new System.Drawing.Point(70, 140);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            
            // btnCancel
            this.btnCancel.Location = new System.Drawing.Point(160, 140);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            
            // Form
            this.ClientSize = new System.Drawing.Size(300, 180);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDiameter1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSlope1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtDiameter2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSlope2);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Text = "Thiết lập thông số độ dốc";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void LoadDefaultValues()
        {
            txtDiameter1.Text = "100";
            txtSlope1.Text = "0.02";
            txtDiameter2.Text = "50";
            txtSlope2.Text = "0.01";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (ValidateInput())
            {
                Diameter1 = double.Parse(txtDiameter1.Text);
                Slope1 = double.Parse(txtSlope1.Text);
                Diameter2 = double.Parse(txtDiameter2.Text);
                Slope2 = double.Parse(txtSlope2.Text);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private bool ValidateInput()
        {
            if (!double.TryParse(txtDiameter1.Text, out double d1) || d1 <= 0)
            {
                MessageBox.Show("Đường kính 1 phải là số dương");
                return false;
            }
            
            if (!double.TryParse(txtSlope1.Text, out double s1) || s1 <= 0)
            {
                MessageBox.Show("Độ dốc 1 phải là số dương");
                return false;
            }
            
            if (!double.TryParse(txtDiameter2.Text, out double d2) || d2 <= 0)
            {
                MessageBox.Show("Đường kính 2 phải là số dương");
                return false;
            }
            
            if (!double.TryParse(txtSlope2.Text, out double s2) || s2 <= 0)
            {
                MessageBox.Show("Độ dốc 2 phải là số dương");
                return false;
            }
            
            if (d1 <= d2)
            {
                MessageBox.Show("Đường kính 1 phải lớn hơn đường kính 2");
                return false;
            }

            return true;
        }
    }
}