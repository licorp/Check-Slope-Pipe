using System;
using System.Collections.Generic;
using System.Linq;
using SWF = System.Windows.Forms;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;
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
                SlopeParametersWindow window = new SlopeParametersWindow(doc);
                bool? result = window.ShowDialog();
                if (result == true)
                {
                    Debug.WriteLine($"[SlopeCheck] Starting pipe slope check with {selectedPipes.Count} pipes");
                    Debug.WriteLine($"[SlopeCheck] Tolerance: {window.Tolerance}%");
                    
                    // Kiểm tra độ dốc cho các đường ống được chọn
                    List<PipeInfo> nonCompliantPipes = CheckPipeSlopes(selectedPipes, window.SizeSlopePairs, window.Tolerance);
                    
                    Debug.WriteLine($"[SlopeCheck] Found {nonCompliantPipes.Count} non-compliant pipes");
                    
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
        private List<PipeInfo> CheckPipeSlopes(List<Pipe> pipes, List<SizeSlopePair> sizeSlopePairs, double tolerance)
        {
            List<PipeInfo> nonCompliantPipes = new List<PipeInfo>();
            Debug.WriteLine($"[SlopeCheck] Checking {pipes.Count} pipes with tolerance {tolerance}%");

            foreach (Pipe pipe in pipes)
            {
                try
                {
                    // Bỏ qua đường ống thẳng đứng
                    if (IsVerticalPipe(pipe))
                    {
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Skipped (vertical pipe)");
                        continue;
                    }

                    // Lấy đường kính đường ống (chuyển đổi từ feet sang mm)
                    double pipeDiameter = GetPipeDiameter(pipe) * 304.8;
                    double pipeSlope = GetPipeSlope(pipe);

                    Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Diameter:{pipeDiameter:F1}mm, Actual Slope:{pipeSlope:F4}");

                    // Tìm size-slope pair phù hợp
                    SizeSlopePair matchingPair = sizeSlopePairs.FirstOrDefault(pair => 
                        Math.Abs(pipeDiameter - pair.Size) < 5.0); // Dung sai 5mm cho size

                    if (matchingPair != null)
                    {
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Matched with Size:{matchingPair.Size}mm, Required Slope:{matchingPair.Slope:F4}");
                        
                        // Chuyển đổi slope từ % sang decimal nếu cần
                        double requiredSlope = matchingPair.Slope;
                        if (requiredSlope > 1.0) // Nếu slope > 1, coi như là %
                        {
                            requiredSlope = requiredSlope / 100.0;
                        }

                        // Tính toán tolerance theo %
                        double toleranceDecimal = tolerance / 100.0;
                        double slopeDifference = Math.Abs(pipeSlope - requiredSlope);
                        double allowedDifference = requiredSlope * toleranceDecimal;

                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Slope difference:{slopeDifference:F6}, Allowed:{allowedDifference:F6}");

                        // Kiểm tra độ dốc với tolerance
                        if (slopeDifference > allowedDifference)
                        {
                            nonCompliantPipes.Add(new PipeInfo
                            {
                                Pipe = pipe,
                                ActualSlope = pipeSlope,
                                RequiredSlope = requiredSlope,
                                Size = pipeDiameter
                            });
                            Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - NON-COMPLIANT");
                        }
                        else
                        {
                            Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - COMPLIANT");
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - No matching size-slope pair found");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SlopeCheck] Error checking pipe ID:{pipe.Id} - {ex.Message}");
                }
            }

            Debug.WriteLine($"[SlopeCheck] Check completed. {nonCompliantPipes.Count} non-compliant pipes found");
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
            try
            {
                Parameter diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (diameterParam != null && diameterParam.HasValue)
                {
                    double diameter = diameterParam.AsDouble();
                    Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Diameter: {diameter:F6} feet = {diameter * 304.8:F1} mm");
                    return diameter;
                }
                Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Unable to get diameter, returning 0");
                return 0.0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlopeCheck] Error getting diameter for pipe ID:{pipe.Id} - {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Lấy độ dốc của đường ống
        /// </summary>
        private double GetPipeSlope(Pipe pipe)
        {
            try
            {
                // Thử lấy từ parameter Slope trước
                Parameter slopeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                if (slopeParam != null && slopeParam.HasValue)
                {
                    double slope = slopeParam.AsDouble();
                    Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Slope from parameter: {slope:F6}");
                    return slope;
                }

                // Nếu không có parameter, tính toán từ geometry
                LocationCurve location = pipe.Location as LocationCurve;
                if (location != null)
                {
                    Curve curve = location.Curve;
                    XYZ startPoint = curve.GetEndPoint(0);
                    XYZ endPoint = curve.GetEndPoint(1);
                    
                    double horizontalDistance = Math.Sqrt(
                        Math.Pow(endPoint.X - startPoint.X, 2) + 
                        Math.Pow(endPoint.Y - startPoint.Y, 2));
                    
                    double verticalDistance = endPoint.Z - startPoint.Z;
                    
                    if (horizontalDistance > 0)
                    {
                        double calculatedSlope = Math.Abs(verticalDistance / horizontalDistance);
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Calculated slope from geometry: {calculatedSlope:F6}");
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Horizontal: {horizontalDistance:F3}, Vertical: {verticalDistance:F3}");
                        return calculatedSlope;
                    }
                }

                Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Unable to determine slope, returning 0");
                return 0.0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlopeCheck] Error getting slope for pipe ID:{pipe.Id} - {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Tạo schedule trong Revit với thông tin chi tiết
        /// </summary>
        private void CreateSchedule(Document doc, List<PipeInfo> nonCompliantPipes)
        {
            using (Transaction trans = new Transaction(doc, "Create/Update Pipe Slope Schedule"))
            {
                trans.Start();

                try
                {
                    Debug.WriteLine("[SlopeCheck] Starting schedule creation/update");
                    
                    // Tạo các tham số tạm thời nếu chưa tồn tại
                    CreateTemporaryParameters(doc);
                    
                    // Đặt giá trị tham số cho các đường ống không đạt yêu cầu
                    SetPipeParameters(doc, nonCompliantPipes);
                    
                    // Kiểm tra và xóa schedule cũ nếu tồn tại
                    ViewSchedule existingSchedule = GetExistingSchedule(doc, "Pipe Slope Check");
                    if (existingSchedule != null)
                    {
                        Debug.WriteLine("[SlopeCheck] Found existing schedule, deleting it");
                        doc.Delete(existingSchedule.Id);
                    }
                    
                    // Tạo schedule mới
                    ViewSchedule schedule = CreatePipeSchedule(doc);
                    Debug.WriteLine($"[SlopeCheck] Created new schedule with ID: {schedule.Id}");
                    
                    // Thêm các trường vào schedule
                    AddScheduleFields(schedule, doc);
                    
                    // Thêm bộ lọc để chỉ hiển thị đường ống không đạt yêu cầu
                    AddScheduleFilters(schedule);
                    
                    trans.Commit();
                    Debug.WriteLine("[SlopeCheck] Schedule creation/update completed successfully");
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    Debug.WriteLine($"[SlopeCheck] Error creating schedule: {ex.Message}");
                    throw new Exception($"Lỗi khi tạo schedule: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Tìm schedule đã tồn tại theo tên
        /// </summary>
        private ViewSchedule GetExistingSchedule(Document doc, string scheduleName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            return collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(schedule => schedule.Name.Equals(scheduleName, StringComparison.OrdinalIgnoreCase));
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
                CategorySet categorySet = new CategorySet();
                categorySet.Insert(pipeCategory);
                Binding binding = new InstanceBinding(categorySet);
                bindingMap.Insert(nonCompliantDef, binding);
            }
            
            if (!bindingMap.Contains(requiredSlopeDef))
            {
                CategorySet categorySet = new CategorySet();
                categorySet.Insert(pipeCategory);
                Binding binding = new InstanceBinding(categorySet);
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
                ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paramName, paramType);
                definition = group.Definitions.Create(options);
            }
            
            return definition;
        }

        /// <summary>
        /// Đặt giá trị tham số cho đường ống
        /// </summary>
        private void SetPipeParameters(Document doc, List<PipeInfo> nonCompliantPipes)
        {
            Debug.WriteLine($"[SlopeCheck] Setting parameters for {nonCompliantPipes.Count} non-compliant pipes");
            
            foreach (PipeInfo pipeInfo in nonCompliantPipes)
            {
                try
                {
                    Pipe pipe = pipeInfo.Pipe;
                    Debug.WriteLine($"[SlopeCheck] Setting parameters for Pipe ID:{pipe.Id}");
                    
                    // Đặt tham số Non-Compliant thành true
                    Parameter nonCompliantParam = pipe.LookupParameter("SlopeCheck_NonCompliant");
                    if (nonCompliantParam != null && !nonCompliantParam.IsReadOnly)
                    {
                        nonCompliantParam.Set(1); // 1 = true cho Yes/No
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Set NonCompliant = true");
                    }
                    else
                    {
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - NonCompliant parameter not found or read-only");
                    }
                    
                    // Đặt tham số Required Slope
                    Parameter requiredSlopeParam = pipe.LookupParameter("SlopeCheck_RequiredSlope");
                    if (requiredSlopeParam != null && !requiredSlopeParam.IsReadOnly)
                    {
                        requiredSlopeParam.Set(pipeInfo.RequiredSlope);
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - Set RequiredSlope = {pipeInfo.RequiredSlope:F4}");
                    }
                    else
                    {
                        Debug.WriteLine($"[SlopeCheck] Pipe ID:{pipe.Id} - RequiredSlope parameter not found or read-only");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SlopeCheck] Error setting parameters for pipe ID:{pipeInfo.Pipe.Id} - {ex.Message}");
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
                f.ParameterId == new ElementId(parameter));
            
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
            for (int i = existingFilters.Count - 1; i >= 0; i--)
            {
                definition.RemoveFilter(i);
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
}