using System;
using System.Collections.Generic;
using System.Linq;
using SWF = System.Windows.Forms;
using System.IO;
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Yêu cầu người dùng quét chọn đường ống
                IList<Reference> selectedReferences;
                try
                {
                    selectedReferences = uiDoc.Selection.PickObjects(ObjectType.Element, 
                        new PipeSelectionFilter(), "Chọn các đường ống cần kiểm tra độ dốc");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                // Lấy các đường ống được chọn
                List<Pipe> selectedPipes = new List<Pipe>();
                foreach (Reference reference in selectedReferences)
                {
                    Pipe pipe = doc.GetElement(reference) as Pipe;
                    if (pipe != null)
                    {
                        selectedPipes.Add(pipe);
                    }
                }

                if (selectedPipes.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Không có đường ống nào được chọn.");
                    return Result.Cancelled;
                }

                // Hiển thị form nhập thông số độ dốc
                SlopeParametersForm form = new SlopeParametersForm();
                if (form.ShowDialog() == SWF.DialogResult.OK)
                {
                    // Kiểm tra độ dốc cho các đường ống được chọn
                    List<PipeInfo> nonCompliantPipes = CheckPipeSlopes(selectedPipes, form.SizeSlopePairs);
                    
                    // Tạo schedule cho các đường ống không đạt yêu cầu
                    if (nonCompliantPipes.Count > 0)
                    {
                        CreateSchedule(doc, nonCompliantPipes);
                        TaskDialog.Show("Kết quả", 
                            $"Đã kiểm tra {selectedPipes.Count} đường ống.\n" +
                            $"Tìm thấy {nonCompliantPipes.Count} đường ống không đạt yêu cầu.\n" +
                            $"Đã tạo schedule 'Pipe Slope Check'.");
                    }
                    else
                    {
                        TaskDialog.Show("Kết quả", "Tất cả đường ống đều đạt yêu cầu về độ dốc.");
                    }
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
        /// Lọc chỉ chọn đường ống
        /// </summary>
        public class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Pipe;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        /// <summary>
        /// Kiểm tra độ dốc của đường ống theo danh sách size-slope
        /// </summary>
        private List<PipeInfo> CheckPipeSlopes(List<Pipe> pipes, List<SizeSlopePair> sizeSlopePairs)
        {
            List<PipeInfo> nonCompliantPipes = new List<PipeInfo>();

            foreach (Pipe pipe in pipes)
            {
                // Bỏ qua đường ống thẳng đứng
                if (IsVerticalPipe(pipe))
                    continue;

                // Lấy đường kính đường ống (chuyển đổi từ feet sang mm)
                double pipeDiameter = GetPipeDiameter(pipe) * 304.8;
                double pipeSlope = GetPipeSlope(pipe);

                // Tìm size-slope pair phù hợp
                SizeSlopePair matchingPair = sizeSlopePairs.FirstOrDefault(pair => 
                    Math.Abs(pipeDiameter - pair.Size) < 0.1); // Dung sai 0.1mm

                if (matchingPair != null)
                {
                    // Kiểm tra độ dốc
                    if (Math.Abs(pipeSlope - matchingPair.Slope) > 0.001) // Dung sai nhỏ
                    {
                        nonCompliantPipes.Add(new PipeInfo
                        {
                            Pipe = pipe,
                            ActualSlope = pipeSlope,
                            RequiredSlope = matchingPair.Slope,
                            Size = pipeDiameter
                        });
                    }
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

            double deltaZ = Math.Abs(endPoint.Z - startPoint.Z);
            double deltaXY = Math.Sqrt(Math.Pow(endPoint.X - startPoint.X, 2) + Math.Pow(endPoint.Y - startPoint.Y, 2));

            return deltaZ > deltaXY * 10;
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
        /// Tạo schedule trong Revit với thông tin chi tiết
        /// </summary>
        private void CreateSchedule(Document doc, List<PipeInfo> nonCompliantPipes)
        {
            using (Transaction trans = new Transaction(doc, "Create Pipe Slope Schedule"))
            {
                trans.Start();

                try
                {
                    // Tạo các tham số tạm thời nếu chưa tồn tại
                    CreateTemporaryParameters(doc);
                    
                    // Đặt giá trị tham số cho các đường ống không đạt yêu cầu
                    SetPipeParameters(doc, nonCompliantPipes);
                    
                    // Tạo schedule mới
                    ViewSchedule schedule = CreatePipeSchedule(doc);
                    
                    // Thêm các trường vào schedule
                    AddScheduleFields(schedule, doc);
                    
                    // Thêm bộ lọc để chỉ hiển thị đường ống không đạt yêu cầu
                    AddScheduleFilters(schedule);
                    
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new Exception($"Lỗi khi tạo schedule: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Tạo các tham số tạm thời cho kiểm tra độ dốc
        /// </summary>
        private void CreateTemporaryParameters(Document doc)
        {
            // Kiểm tra xem tham số đã tồn tại chưa
            Category pipeCategory = Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves);
            BindingMap bindingMap = doc.ParameterBindings;
            
            // Tạo định nghĩa tham số cho "Non-Compliant"
            string nonCompliantParamName = "SlopeCheck_NonCompliant";
            DefinitionFile defFile = GetOrCreateSharedParameterFile(doc);
            Definition nonCompliantDef = GetOrCreateDefinition(defFile, nonCompliantParamName, SpecTypeId.Boolean.YesNo);
            
            // Tạo định nghĩa tham số cho "Required Slope"
            string requiredSlopeParamName = "SlopeCheck_RequiredSlope";
            Definition requiredSlopeDef = GetOrCreateDefinition(defFile, requiredSlopeParamName, SpecTypeId.Number);
            
            // Tạo binding nếu chưa tồn tại
            if (!bindingMap.Contains(nonCompliantDef))
            {
                Binding binding = new InstanceBinding(new List<Category> { pipeCategory });
                bindingMap.Insert(nonCompliantDef, binding);
            }
            
            if (!bindingMap.Contains(requiredSlopeDef))
            {
                Binding binding = new InstanceBinding(new List<Category> { pipeCategory });
                bindingMap.Insert(requiredSlopeDef, binding);
            }
        }

        /// <summary>
        /// Lấy hoặc tạo file tham số chia sẻ
        /// </summary>
        private DefinitionFile GetOrCreateSharedParameterFile(Document doc)
        {
            Application app = doc.Application;
            string tempPath = System.IO.Path.GetTempPath();
            string sharedParamFile = System.IO.Path.Combine(tempPath, "SlopeCheck_SharedParameters.txt");
            
            // Tạo file tham số chia sẻ tạm thời
            if (!System.IO.File.Exists(sharedParamFile))
            {
                System.IO.File.WriteAllText(sharedParamFile, "");
            }
            
            app.SharedParametersFilename = sharedParamFile;
            return app.OpenSharedParameterFile();
        }

        /// <summary>
        /// Lấy hoặc tạo định nghĩa tham số
        /// </summary>
        private Definition GetOrCreateDefinition(DefinitionFile defFile, string paramName, ForgeTypeId paramType)
        {
            Definition definition = defFile.Groups
                .SelectMany(g => g.Definitions)
                .FirstOrDefault(d => d.Name == paramName);
            
            if (definition == null)
            {
                DefinitionGroup group = defFile.Groups.FirstOrDefault() ?? defFile.Groups.Create("SlopeCheck");
                definition = group.Definitions.Create(paramName, paramType);
            }
            
            return definition;
        }

        /// <summary>
        /// Đặt giá trị tham số cho đường ống
        /// </summary>
        private void SetPipeParameters(Document doc, List<PipeInfo> nonCompliantPipes)
        {
            foreach (PipeInfo pipeInfo in nonCompliantPipes)
            {
                Pipe pipe = pipeInfo.Pipe;
                
                // Đặt tham số Non-Compliant thành true
                Parameter nonCompliantParam = pipe.LookupParameter("SlopeCheck_NonCompliant");
                if (nonCompliantParam != null && !nonCompliantParam.IsReadOnly)
                {
                    nonCompliantParam.Set(1); // 1 = true cho Yes/No
                }
                
                // Đặt tham số Required Slope
                Parameter requiredSlopeParam = pipe.LookupParameter("SlopeCheck_RequiredSlope");
                if (requiredSlopeParam != null && !requiredSlopeParam.IsReadOnly)
                {
                    requiredSlopeParam.Set(pipeInfo.RequiredSlope);
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
            schedule.Name = "Pipe Slope Check - Non Compliant";
            
            return schedule;
        }

        /// <summary>
        /// Thêm các trường vào schedule
        /// </summary>
        private void AddScheduleFields(ViewSchedule schedule, Document doc)
        {
            ScheduleDefinition definition = schedule.Definition;
            
            // Xóa các trường mặc định
            IList<ScheduleFieldId> existingFields = definition.GetFieldOrder();
            foreach (ScheduleFieldId fieldId in existingFields.ToList())
            {
                definition.RemoveField(fieldId);
            }

            // Thêm các trường theo yêu cầu
            // 1. ID element
            AddScheduleFieldByParameter(definition, BuiltInParameter.ID_PARAM, "Element ID");
            
            // 2. Size (Diameter)
            AddScheduleFieldByParameter(definition, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, "Size");
            
            // 3. Slope thực tế
            AddScheduleFieldByParameter(definition, BuiltInParameter.RBS_PIPE_SLOPE, "Actual Slope");
            
            // 4. Slope yêu cầu (tham số tạm thời)
            AddScheduleFieldByName(definition, "SlopeCheck_RequiredSlope", "Required Slope", doc);
            
            // 5. Trạng thái không đạt yêu cầu
            AddScheduleFieldByName(definition, "SlopeCheck_NonCompliant", "Non-Compliant", doc);
        }

        /// <summary>
        /// Thêm trường schedule dựa trên parameter
        /// </summary>
        private void AddScheduleFieldByParameter(ScheduleDefinition definition, BuiltInParameter parameter, string columnName)
        {
            // Lấy danh sách các trường có thể lập lịch
            IList<SchedulableField> schedulableFields = definition.GetSchedulableFields();
            
            // Tìm trường phù hợp với parameter
            SchedulableField targetField = schedulableFields.FirstOrDefault(f =>
                f.ParameterId == parameter);
            
            if (targetField != null)
            {
                ScheduleField field = definition.AddField(targetField);
                field.ColumnHeading = columnName;
            }
        }

        /// <summary>
        /// Thêm trường schedule dựa trên tên tham số
        /// </summary>
        private void AddScheduleFieldByName(ScheduleDefinition definition, string parameterName, string columnName, Document doc)
        {
            IList<SchedulableField> schedulableFields = definition.GetSchedulableFields();
            
            // Tìm trường bằng tên
            SchedulableField targetField = schedulableFields.FirstOrDefault(f =>
            {
                string name = f.GetName(doc);
                return name.Equals(parameterName, StringComparison.OrdinalIgnoreCase);
            });
            
            if (targetField != null)
            {
                ScheduleField field = definition.AddField(targetField);
                field.ColumnHeading = columnName;
            }
        }

        /// <summary>
        /// Thêm bộ lọc schedule để chỉ hiển thị đường ống không đạt yêu cầu
        /// </summary>
        private void AddScheduleFilters(ViewSchedule schedule)
        {
            ScheduleDefinition definition = schedule.Definition;
            
            // Xóa các bộ lọc hiện có
            IList<ScheduleFilter> existingFilters = definition.GetFilters();
            foreach (ScheduleFilter filter in existingFilters)
            {
                definition.RemoveFilter(filter);
            }

            // Tìm trường Non-Compliant
            ScheduleField nonCompliantField = definition.GetFieldOrder()
                .Select(fieldId => definition.GetField(fieldId))
                .FirstOrDefault(f => f.ColumnHeading == "Non-Compliant");
            
            if (nonCompliantField != null)
            {
                // Tạo bộ lọc: Non-Compliant = true
                ScheduleFilter filter = new ScheduleFilter(nonCompliantField.FieldId, ScheduleFilterType.Equal, 1);
                definition.AddFilter(filter);
            }
        }
    }

    /// <summary>
    /// Class lưu thông tin đường ống không đạt yêu cầu
    /// </summary>
    public class PipeInfo
    {
        public Pipe Pipe { get; set; }
        public double Size { get; set; }
        public double ActualSlope { get; set; }
        public double RequiredSlope { get; set; }
    }

    /// <summary>
    /// Class lưu cặp size-slope
    /// </summary>
    public class SizeSlopePair
    {
        public double Size { get; set; }
        public double Slope { get; set; }
    }

    /// <summary>
    /// Form nhập thông số độ dốc với khả năng thêm nhiều cặp size-slope
    /// </summary>
    public partial class SlopeParametersForm : SWF.Form
    {
        public List<SizeSlopePair> SizeSlopePairs { get; private set; }
        
        private SWF.FlowLayoutPanel flowPanel;
        private SWF.Button btnAdd;
        private SWF.Button btnOK;
        private SWF.Button btnCancel;
        private int rowCount = 0;

        public SlopeParametersForm()
        {
            SizeSlopePairs = new List<SizeSlopePair>();
            InitializeComponent();
            AddInitialRow();
        }

        private void InitializeComponent()
        {
            this.Text = "Thiết lập thông số độ dốc";
            this.Size = new System.Drawing.Size(400, 300);
            this.StartPosition = SWF.FormStartPosition.CenterScreen;
            this.FormBorderStyle = SWF.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Tạo flow layout panel
            flowPanel = new SWF.FlowLayoutPanel();
            flowPanel.FlowDirection = SWF.FlowDirection.TopDown;
            flowPanel.AutoScroll = true;
            flowPanel.Location = new System.Drawing.Point(10, 10);
            flowPanel.Size = new System.Drawing.Size(370, 200);
            this.Controls.Add(flowPanel);

            // Tạo nút thêm
            btnAdd = new SWF.Button();
            btnAdd.Text = "+";
            btnAdd.Size = new System.Drawing.Size(30, 23);
            btnAdd.Location = new System.Drawing.Point(10, 220);
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);

            // Tạo nút OK
            btnOK = new SWF.Button();
            btnOK.Text = "OK";
            btnOK.Size = new System.Drawing.Size(75, 23);
            btnOK.Location = new System.Drawing.Point(200, 250);
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Tạo nút Cancel
            btnCancel = new SWF.Button();
            btnCancel.Text = "Cancel";
            btnCancel.Size = new System.Drawing.Size(75, 23);
            btnCancel.Location = new System.Drawing.Point(285, 250);
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
        }

        private void AddInitialRow()
        {
            AddNewRow();
        }

        private void AddNewRow()
        {
            rowCount++;
            
            // Tạo panel cho dòng mới
            SWF.Panel rowPanel = new SWF.Panel();
            rowPanel.Size = new System.Drawing.Size(350, 30);
            rowPanel.Margin = new SWF.Padding(0, 5, 0, 5);

            // Label size
            SWF.Label lblSize = new SWF.Label();
            lblSize.Text = $"Size {rowCount} (mm):";
            lblSize.Location = new System.Drawing.Point(0, 5);
            lblSize.Size = new System.Drawing.Size(80, 20);
            rowPanel.Controls.Add(lblSize);

            // Textbox size
            SWF.TextBox txtSize = new SWF.TextBox();
            txtSize.Location = new System.Drawing.Point(85, 5);
            txtSize.Size = new System.Drawing.Size(60, 20);
            txtSize.Tag = rowCount;
            rowPanel.Controls.Add(txtSize);

            // Label slope
            SWF.Label lblSlope = new SWF.Label();
            lblSlope.Text = "Slope:";
            lblSlope.Location = new System.Drawing.Point(155, 5);
            lblSlope.Size = new System.Drawing.Size(40, 20);
            rowPanel.Controls.Add(lblSlope);

            // Textbox slope
            SWF.TextBox txtSlope = new SWF.TextBox();
            txtSlope.Location = new System.Drawing.Point(200, 5);
            txtSlope.Size = new System.Drawing.Size(60, 20);
            txtSlope.Tag = rowCount;
            rowPanel.Controls.Add(txtSlope);

            // Nút xóa
            SWF.Button btnRemove = new SWF.Button();
            btnRemove.Text = "X";
            btnRemove.Size = new System.Drawing.Size(25, 20);
            btnRemove.Location = new System.Drawing.Point(270, 5);
            btnRemove.Tag = rowCount;
            btnRemove.Click += BtnRemove_Click;
            rowPanel.Controls.Add(btnRemove);

            flowPanel.Controls.Add(rowPanel);
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            AddNewRow();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                int rowNumber = (int)btn.Tag;
                SWF.Control rowPanel = flowPanel.Controls
                    .OfType<SWF.Panel>()
                    .FirstOrDefault(p => p.Controls.OfType<SWF.TextBox>()
                        .Any(txt => (int?)txt.Tag == rowNumber));
                
                if (rowPanel != null)
                {
                    flowPanel.Controls.Remove(rowPanel);
                    // Cập nhật lại số thứ tự các dòng còn lại
                    UpdateRowNumbers();
                }
            }
        }

        private void UpdateRowNumbers()
        {
            var rows = flowPanel.Controls.OfType<SWF.Panel>().ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                var labels = rows[i].Controls.OfType<SWF.Label>().Where(l => l.Text.StartsWith("Size")).ToList();
                if (labels.Count > 0)
                {
                    labels[0].Text = $"Size {i + 1} (mm):";
                }
                
                var textBoxes = rows[i].Controls.OfType<SWF.TextBox>().ToList();
                foreach (var txt in textBoxes)
                {
                    txt.Tag = i + 1;
                }
                
                var buttons = rows[i].Controls.OfType<SWF.Button>().ToList();
                foreach (var btn in buttons)
                {
                    btn.Tag = i + 1;
                }
            }
            rowCount = rows.Count;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (ValidateInput())
            {
                SizeSlopePairs.Clear();
                
                // Thu thập dữ liệu từ các dòng
                foreach (SWF.Panel rowPanel in flowPanel.Controls.OfType<SWF.Panel>())
                {
                    var textBoxes = rowPanel.Controls.OfType<SWF.TextBox>().ToList();
                    if (textBoxes.Count >= 2)
                    {
                        double size = double.Parse(textBoxes[0].Text);
                        double slope = double.Parse(textBoxes[1].Text);
                        
                        SizeSlopePairs.Add(new SizeSlopePair { Size = size, Slope = slope });
                    }
                }
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private bool ValidateInput()
        {
            // Kiểm tra từng dòng
            foreach (SWF.Panel rowPanel in flowPanel.Controls.OfType<SWF.Panel>())
            {
                var textBoxes = rowPanel.Controls.OfType<SWF.TextBox>().ToList();
                if (textBoxes.Count < 2) continue;

                if (!double.TryParse(textBoxes[0].Text, out double size) || size <= 0)
                {
                    SWF.MessageBox.Show("Size phải là số dương", "Lỗi", SWF.MessageBoxButtons.OK, SWF.MessageBoxIcon.Error);
                    textBoxes[0].Focus();
                    return false;
                }

                if (!double.TryParse(textBoxes[1].Text, out double slope) || slope <= 0)
                {
                    SWF.MessageBox.Show("Slope phải là số dương", "Lỗi", SWF.MessageBoxButtons.OK, SWF.MessageBoxIcon.Error);
                    textBoxes[1].Focus();
                    return false;
                }
            }

            if (flowPanel.Controls.OfType<SWF.Panel>().Count() == 0)
            {
                SWF.MessageBox.Show("Cần ít nhất một cặp size-slope", "Lỗi", SWF.MessageBoxButtons.OK, SWF.MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
    }
}