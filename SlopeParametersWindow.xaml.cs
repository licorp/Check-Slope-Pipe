using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;
using Grid = System.Windows.Controls.Grid;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace CheckSlopePipe
{
    public partial class SlopeParametersWindow : Window
    {
        public List<SizeSlopePair> SizeSlopePairs { get; private set; }
        public double Tolerance { get; private set; }
        private Document RevitDocument { get; set; }

        // Available sizes and slopes from model
        private List<double> AvailableSizes = new List<double>();
        private List<double> AvailableSlopes = new List<double>();

        // Predefined size options (common pipe sizes in mm) - fallback
        private readonly List<double> PredefinedSizes = new List<double>
        {
            50, 75, 100, 125, 150, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800, 900, 1000
        };

        // Predefined slope options (in %) - fallback
        private readonly List<double> PredefinedSlopes = new List<double>
        {
            0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0
        };
        
        public SlopeParametersWindow(Document doc)
        {
            InitializeComponent();
            SizeSlopePairs = new List<SizeSlopePair>();
            RevitDocument = doc;
            
            // Get actual sizes and slopes from model
            GetAvailableSizesAndSlopes();
            
            // Add default rows with actual data from model
            if (AvailableSizes.Count > 0)
            {
                for (int i = 0; i < Math.Min(3, AvailableSizes.Count); i++)
                {
                    double slope = i < AvailableSlopes.Count ? AvailableSlopes[i] : AvailableSlopes[0];
                    AddSizeSlopeRow(AvailableSizes[i], slope);
                }
            }
            else
            {
                // Fallback to predefined values
                AddSizeSlopeRow(PredefinedSizes[2], PredefinedSlopes[1]); // 100mm, 1.0%
                AddSizeSlopeRow(PredefinedSizes[4], PredefinedSlopes[2]); // 150mm, 1.5%
                AddSizeSlopeRow(PredefinedSizes[5], PredefinedSlopes[3]); // 200mm, 2.0%
            }
            
            Debug.WriteLine($"[SlopeCheck] SlopeParametersWindow initialized with {AvailableSizes.Count} sizes and {AvailableSlopes.Count} slopes from model");
        }

        private void GetAvailableSizesAndSlopes()
        {
            AvailableSizes = new List<double>();
            AvailableSlopes = new List<double>();
            
            try
            {
                // Get all pipes in the model
                FilteredElementCollector collector = new FilteredElementCollector(RevitDocument);
                ICollection<Element> pipes = collector.OfClass(typeof(Pipe)).ToElements();
                
                Debug.WriteLine($"[SlopeCheck] Found {pipes.Count} pipes in model");

                foreach (Pipe pipe in pipes.Cast<Pipe>())
                {
                    try
                    {
                        // Get diameter
                        Parameter diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParam != null && diameterParam.HasValue)
                        {
                            double diameterFeet = diameterParam.AsDouble();
                            double diameterMm = diameterFeet * 304.8;
                            
                            // Round to nearest mm and add if not already in list
                            diameterMm = Math.Round(diameterMm, 0);
                            if (!AvailableSizes.Contains(diameterMm) && diameterMm > 0)
                            {
                                AvailableSizes.Add(diameterMm);
                            }
                        }

                        // Get slope if available
                        Parameter slopeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                        if (slopeParam != null && slopeParam.HasValue)
                        {
                            double slope = slopeParam.AsDouble();
                            if (slope > 0)
                            {
                                // Convert to percentage if it's a decimal
                                if (slope < 1)
                                {
                                    slope = slope * 100;
                                }
                                slope = Math.Round(slope, 2);
                                if (!AvailableSlopes.Contains(slope))
                                {
                                    AvailableSlopes.Add(slope);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[SlopeCheck] Error processing pipe {pipe.Id}: {ex.Message}");
                    }
                }
                
                // Sort the lists
                AvailableSizes.Sort();
                AvailableSlopes.Sort();
                
                // If no slopes found, add common values
                if (AvailableSlopes.Count == 0)
                {
                    AvailableSlopes.AddRange(PredefinedSlopes);
                }
                
                // If no sizes found, add common sizes
                if (AvailableSizes.Count == 0)
                {
                    AvailableSizes.AddRange(PredefinedSizes);
                }

                Debug.WriteLine($"[SlopeCheck] Available sizes: {string.Join(", ", AvailableSizes)}");
                Debug.WriteLine($"[SlopeCheck] Available slopes: {string.Join(", ", AvailableSlopes)}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlopeCheck] Error getting sizes/slopes from model: {ex.Message}");
                
                // Use predefined values if error
                AvailableSizes = new List<double>(PredefinedSizes);
                AvailableSlopes = new List<double>(PredefinedSlopes);
            }
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            AddSizeSlopeRow(100, 1.0);
            Debug.WriteLine("[SlopeCheck] New row added");
        }

        private void AddSizeSlopeRow(double selectedSize = 100, double selectedSlope = 1.0)
        {
            Grid rowGrid = new Grid();
            rowGrid.Height = 40; // Increased height for ComboBox
            rowGrid.Margin = new Thickness(0, 2, 0, 2);
            rowGrid.Background = System.Windows.Media.Brushes.White;
            
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Size ComboBox
            ComboBox sizeComboBox = new ComboBox
            {
                Margin = new Thickness(5, 5, 5, 5),
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                IsEditable = true,
                VerticalAlignment = VerticalAlignment.Center
            };
            
            // Add available sizes from model or predefined
            var sizesToUse = AvailableSizes.Count > 0 ? AvailableSizes : PredefinedSizes;
            foreach (double size in sizesToUse)
            {
                sizeComboBox.Items.Add(size.ToString("0"));
            }
            
            sizeComboBox.Text = selectedSize.ToString("0");
            Grid.SetColumn(sizeComboBox, 0);
            rowGrid.Children.Add(sizeComboBox);

            // Slope ComboBox
            ComboBox slopeComboBox = new ComboBox
            {
                Margin = new Thickness(5, 5, 5, 5),
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                IsEditable = true,
                VerticalAlignment = VerticalAlignment.Center
            };
            
            // Add available slopes from model or predefined
            var slopesToUse = AvailableSlopes.Count > 0 ? AvailableSlopes : PredefinedSlopes;
            foreach (double slope in slopesToUse)
            {
                slopeComboBox.Items.Add(slope.ToString("0.0"));
            }
            
            slopeComboBox.Text = selectedSlope.ToString("0.0");
            Grid.SetColumn(slopeComboBox, 1);
            rowGrid.Children.Add(slopeComboBox);

            // Remove Button
            Button removeButton = new Button
            {
                Content = "✕",
                Width = 30,
                Height = 30,
                Background = System.Windows.Media.Brushes.Red,
                Foreground = System.Windows.Media.Brushes.White,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 12
            };
            removeButton.Click += (s, args) =>
            {
                SizeSlopePanel.Children.Remove(rowGrid);
                Debug.WriteLine("[SlopeCheck] Row removed");
            };
            Grid.SetColumn(removeButton, 2);
            rowGrid.Children.Add(removeButton);

            SizeSlopePanel.Children.Add(rowGrid);
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateInput())
            {
                CollectData();
                DialogResult = true;
                Debug.WriteLine($"[SlopeCheck] Run clicked - {SizeSlopePairs.Count} size-slope pairs, Tolerance: {Tolerance}%");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Debug.WriteLine("[SlopeCheck] Cancelled by user");
        }

        private bool ValidateInput()
        {
            // Validate tolerance
            if (!double.TryParse(ToleranceTextBox.Text, out double tolerance) || tolerance < 0 || tolerance > 100)
            {
                MessageBox.Show("Tolerance must be a number between 0 and 100 / Dung sai phải là số từ 0 đến 100", 
                    "Error / Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                ToleranceTextBox.Focus();
                Debug.WriteLine("[SlopeCheck] Invalid tolerance input");
                return false;
            }

            // Validate size-slope pairs
            foreach (Grid rowGrid in SizeSlopePanel.Children.OfType<Grid>())
            {
                ComboBox sizeComboBox = rowGrid.Children.OfType<ComboBox>().First();
                ComboBox slopeComboBox = rowGrid.Children.OfType<ComboBox>().Skip(1).First();

                if (!double.TryParse(sizeComboBox.Text, out double size) || size <= 0)
                {
                    MessageBox.Show("Size must be a positive number / Size phải là số dương", 
                        "Error / Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    sizeComboBox.Focus();
                    Debug.WriteLine("[SlopeCheck] Invalid size input");
                    return false;
                }

                if (!double.TryParse(slopeComboBox.Text, out double slope) || slope <= 0)
                {
                    MessageBox.Show("Slope must be a positive number / Slope phải là số dương", 
                        "Error / Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    slopeComboBox.Focus();
                    Debug.WriteLine("[SlopeCheck] Invalid slope input");
                    return false;
                }
            }

            if (SizeSlopePanel.Children.Count == 0)
            {
                MessageBox.Show("At least one size-slope pair is required / Cần ít nhất một cặp size-slope", 
                    "Error / Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine("[SlopeCheck] No size-slope pairs defined");
                return false;
            }

            return true;
        }

        private void CollectData()
        {
            // Get tolerance
            Tolerance = double.Parse(ToleranceTextBox.Text);

            // Collect size-slope pairs
            SizeSlopePairs.Clear();
            foreach (Grid rowGrid in SizeSlopePanel.Children.OfType<Grid>())
            {
                ComboBox sizeComboBox = rowGrid.Children.OfType<ComboBox>().First();
                ComboBox slopeComboBox = rowGrid.Children.OfType<ComboBox>().Skip(1).First();

                double size = double.Parse(sizeComboBox.Text);
                double slope = double.Parse(slopeComboBox.Text);

                SizeSlopePairs.Add(new SizeSlopePair { Size = size, Slope = slope });
                Debug.WriteLine($"[SlopeCheck] Added pair: Size={size}mm, Slope={slope}%");
            }
        }
    }
}