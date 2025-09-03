using Microsoft.Win32;
using OfficeOpenXml;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace UPS_OrgChart_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
		private List<OrgChartViewModel> _viewModel;
		/*
		 * A4 Shett page size is 8.2677 inch × 11.6929 inch.
		 * Portrait height ≈ 842 px
		 * Landscape height ≈ 595 px
		 * DPI	-	Dots Per Inch -> For printer
		 * DIP	-	Device Independent Pixel (1/96 inch)
		 * WPF is resolution independent and uses DIPs for layout and rendering.
		 * 96 DPI (WPF DIPs to px)
		 * Portrait height ≈ 1122.52 px
		 * Landscape height ≈ 793.70 px
		 * 300 DPI (your export target)
		 * Portrait height ≈ 3507.87 px
		 * Landscape height ≈ 2480.31 px
		 */
		private static double _a4PageHeight = 1122;
		private static double _a4PageHeightExcludeLegend = _a4PageHeight - 15;
		private const string ItcPod = "";
		public MainWindow()
        {
            InitializeComponent();
			Loaded += MainWindow_Loaded;
		}
		
		private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
		{
		}

		private async void UploadExcelButton_Click(object sender, RoutedEventArgs e)
		{
			var openFileDialog = new OpenFileDialog
			{
				Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
			};
			if (openFileDialog.ShowDialog() == true)
			{
				_viewModel = await LoadOrgChartFromExcel(openFileDialog.FileName);
				DataContext = _viewModel;
				DrawOrgChartHorizontalLayout();
				// DrawOrgChartVerticalLayout();
			}
			else
			{
				MessageBox.Show("No file selected. Application will exit.");
				Application.Current.Shutdown();
			}
		}
		private async Task<List<OrgChartViewModel>> LoadOrgChartFromExcel(string filePath)
		{
			var viewModel = new List<OrgChartViewModel>();
			var nodes = new List<Node>();

			// Example: Excel columns - Id, Name, ParentId, BackgroundColor
			ExcelPackage.License.SetNonCommercialPersonal("LocalTest");
			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets[0];			
				int row = 2;
				while (worksheet.Cells[row, 1].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
				{
					var node = new Node();
					node.ItcAdm = worksheet.Cells[row, 1].Text?.Trim();
					node.OnsiteDirector = worksheet.Cells[row, 2].Text?.Trim();
					node.OnsiteManager = worksheet.Cells[row, 3].Text?.Trim();
					node.RoleTitle = worksheet.Cells[row, 4].Text?.Trim();
					node.PayGrade = worksheet.Cells[row, 5].Text?.Trim();
					node.Status = worksheet.Cells[row, 6].Text?.Trim();
					node.Name = worksheet.Cells[row, 7].Text?.Trim();
					node.Req = worksheet.Cells[row, 8].Text?.Trim();
					nodes.Add(node);
					row++;
				}
			}

			var offShoreAdms = nodes.GroupBy(n => n.ItcAdm);
			foreach (var offShoreAdm in offShoreAdms)
			{
				var offShoreModel = new OrgChartViewModel();
				offShoreModel.OffShoreAdm = offShoreAdm.Key;
				var directors = offShoreAdm.GroupBy(n => n.OnsiteDirector);
				foreach (var subordinates in directors)
				{
					var directorNode = new DirectorNode();
					directorNode.Name = subordinates.Key;
					directorNode.RoleTitle = "Onsite Director";
					directorNode.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffdd99"));

					var managers = subordinates
						.GroupBy(s => string.IsNullOrWhiteSpace(s.OnsiteManager) ? ItcPod : s.OnsiteManager).OrderBy(o => o.Key);
					foreach (var managerSubordinates in managers)
					{
						var managerNode = new ManagerNode();
						managerNode.Name = managerSubordinates.Key;
						managerNode.RoleTitle = "Onsite ADM";
						managerNode.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF9E3"));
						foreach (var member in managerSubordinates.OrderByDescending(o => o.PayGrade))
						{							
							var isOnSeat = string.Equals(member.Status, "On seat - Regular", StringComparison.InvariantCultureIgnoreCase);
							var memberNode = new OrgNode
							{
								Name = isOnSeat ? member.Name : member.Req,
								RoleTitle = member.RoleTitle,
								PayGrade = member.PayGrade,
								Status = member.Status,
								Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(GetColourCode(isOnSeat ? member.PayGrade : GetNonSeatMemberPayGrade(member))))
							};
							managerNode.Employees.Add(memberNode);
						}
						directorNode.Managers.Add(managerNode);
					}
					offShoreModel.DirectorNodes.Add(directorNode);
				}
				viewModel.Add(offShoreModel);
			}

			return viewModel;
		}

		private string GetNonSeatMemberPayGrade(Node member)
		{
			if (member.Status.Equals(OrgChartViewModel.OpenHireAhead, StringComparison.InvariantCultureIgnoreCase))
			{
				return "HIREAHEAD-OPEN";
			}
			if (member.Status.Equals(OrgChartViewModel.OnSeatHireAhead, StringComparison.InvariantCultureIgnoreCase))
			{
				return "HIREAHEAD";
			}
			if (member.Status.Equals(OrgChartViewModel.OfferAccepted, StringComparison.InvariantCultureIgnoreCase))
			{
				return "OFFERED";
			}
			if (member.Status.Equals(OrgChartViewModel.GD, StringComparison.InvariantCultureIgnoreCase))
			{
				return "GD";
			}
			if (member.Status.Equals(OrgChartViewModel.InterviewInProgress, StringComparison.InvariantCultureIgnoreCase))
			{
				return "OPEN";
			}
			return string.Empty;
		}
		private string GetColourCode(string payGrade)
		{
			return payGrade.ToUpper() switch
			{
				"0U" => "#FFFFFF", // White
				"1U" => "#FF69B4", // Hot Pink - #FF69B4, and Light Pink - #FFB6C1 
				"2U" => "#FFFF00", // Yellow
				"HIREAHEAD" => "#DDEEFF", // Light Blue
				"HIREAHEAD-OPEN" => "#E6F2FF",
				"OFFERED" => "#FFD700", // Gold
				"GD" => "#B7E1A1", // Light Green
				"OPEN" => "#A9A9A9", // Dark Gray
				_ => "#FFFFFF"  // Default White
			};
		}
		/*
		private void DrawOrgChartVerticalLayout()
		{
			OrgChartVerticalCanvas.Children.Clear();

			if (_viewModel == null || _viewModel.DirectorNodes == null || _viewModel.DirectorNodes.Count == 0)
				return;

			double startX = 30, startY = 50;
			double directorNodeWidth = 200, directorNodeHeight = 22, directorSpacingX = 40;
			double managerNodeWidth = 160, managerNodeHeight = 28, managerSpacingY = 40;
			double employeeNodeWidth = 160, employeeNodeHeight = 28, employeeSpacingY = 8, employeeSpacingX = 8;
			int batchSize = 8;

			List<Border> directorBorders = new();
			List<double> directorWidths = new();
			List<List<double>> managerWidthsPerDirector = new();

			// Calculate widths for each director's managers and their employee columns
			for (int d = 0; d < _viewModel.DirectorNodes.Count; d++)
			{
				var director = _viewModel.DirectorNodes[d];
				int managerCount = director.Managers.Count;
				double width = 0;
				List<double> managerWidths = new();
				for (int m = 0; m < managerCount; m++)
				{
					var manager = director.Managers[m];
					int employeeCount = manager.Employees?.Count ?? 0;
					int columns = (employeeCount + batchSize - 1) / batchSize;
					double managerWidth = Math.Max(managerNodeWidth, columns * (employeeNodeWidth + employeeSpacingX) - employeeSpacingX);
					managerWidths.Add(managerWidth);
					width += managerWidth + directorSpacingX;
				}
				managerWidthsPerDirector.Add(managerWidths);
				directorWidths.Add(Math.Max(width - directorSpacingX, directorNodeWidth));
			}

			// Layout directors horizontally at the top, centered above their managers
			double currentX = startX;
			for (int d = 0; d < _viewModel.DirectorNodes.Count; d++)
			{
				var director = _viewModel.DirectorNodes[d];
				double directorWidth = directorWidths[d];
				var directorBorder = new Border
				{
					Width = directorNodeWidth,
					Height = directorNodeHeight,
					Background = director.Background ?? Brushes.LightYellow,
					BorderBrush = Brushes.Black,
					BorderThickness = new Thickness(2),
					CornerRadius = new CornerRadius(8),
					Child = new StackPanel
					{
						Margin = new Thickness(4, 0, 0, 0),
						VerticalAlignment = VerticalAlignment.Center,
						HorizontalAlignment = HorizontalAlignment.Center,
						Children =
												{
													new TextBlock
													{
														Text = $"{director.Name}, {director.RoleTitle}",
														FontWeight = FontWeights.Bold,
														FontSize = 8,
														Foreground = Brushes.Black,
														TextAlignment = TextAlignment.Center,
														VerticalAlignment = VerticalAlignment.Center,
														HorizontalAlignment = HorizontalAlignment.Center
													}
												}
					}
				};
				OrgChartVerticalCanvas.Children.Add(directorBorder);
				double directorX = currentX + (directorWidth - directorNodeWidth) / 2;
				Canvas.SetLeft(directorBorder, directorX);
				Canvas.SetTop(directorBorder, startY);
				directorBorders.Add(directorBorder);
				currentX += directorWidth + directorSpacingX;
			}

			// Draw top line and connectors if multiple directors
			if (directorBorders.Count > 1)
			{
				double lineY = startY - 20;
				double leftX = Canvas.GetLeft(directorBorders[0]) + directorNodeWidth / 2;
				double rightX = Canvas.GetLeft(directorBorders[^1]) + directorNodeWidth / 2;
				var topLine = new Line
				{
					X1 = leftX,
					Y1 = lineY,
					X2 = rightX,
					Y2 = lineY,
					Stroke = Brushes.Black,
					StrokeThickness = 2
				};
				OrgChartVerticalCanvas.Children.Add(topLine);

				foreach (var border in directorBorders)
				{
					double midX = Canvas.GetLeft(border) + directorNodeWidth / 2;
					var connector = new Line
					{
						X1 = midX,
						Y1 = lineY,
						X2 = midX,
						Y2 = startY,
						Stroke = Brushes.Black,
						StrokeThickness = 1.5,
						StrokeDashArray = new DoubleCollection { 4, 2 }
					};
					OrgChartVerticalCanvas.Children.Add(connector);
					DrawArrowHead(midX, startY, midX, lineY, OrgChartVerticalCanvas);
				}
			}

			// Draw managers and employees
			for (int d = 0; d < _viewModel.DirectorNodes.Count; d++)
			{
				var director = _viewModel.DirectorNodes[d];
				double directorX = Canvas.GetLeft(directorBorders[d]);
				double directorY = Canvas.GetTop(directorBorders[d]);
				double managerRowY = directorY + directorNodeHeight + 50;

				int managerCount = director.Managers.Count;
				var managerWidths = managerWidthsPerDirector[d];
				double totalManagersWidth = managerWidths.Sum() + directorSpacingX * (managerCount - 1);
				double managerStartX = directorX + (directorNodeWidth / 2) - (totalManagersWidth / 2);

				List<Border> managerBorders = new();

				double nextManagerX = managerStartX;
				for (int m = 0; m < managerCount; m++)
				{
					var manager = director.Managers[m];
					double managerWidth = managerWidths[m];
					int employeeCount = manager.Employees?.Count ?? 0;
					int columns = (employeeCount + batchSize - 1) / batchSize;

					// Center manager node above its employee columns
					double managerX = nextManagerX + (managerWidth - managerNodeWidth) / 2;

					// Draw manager node
					var managerBorder = new Border
					{
						Width = managerNodeWidth,
						Height = managerNodeHeight,
						Background = manager.Background ?? Brushes.LightGoldenrodYellow,
						BorderBrush = Brushes.Black,
						BorderThickness = new Thickness(1.5),
						CornerRadius = new CornerRadius(7),
						Child = new StackPanel
						{
							Margin = new Thickness(4, 0, 0, 0),
							VerticalAlignment = VerticalAlignment.Center,
							HorizontalAlignment = HorizontalAlignment.Center,
							Children =
													{
														new TextBlock
														{
															Text = manager.Name,
															FontWeight = FontWeights.SemiBold,
															FontSize = 8,
															Foreground = Brushes.Black,
															TextAlignment = TextAlignment.Center,
															VerticalAlignment = VerticalAlignment.Center,
															HorizontalAlignment = HorizontalAlignment.Center
														},
														new TextBlock
														{
															Text = manager.RoleTitle,
															FontSize = 8,
															Foreground = Brushes.Black,
															TextAlignment = TextAlignment.Center,
															VerticalAlignment = VerticalAlignment.Center,
															HorizontalAlignment = HorizontalAlignment.Center
														}
													}
						}
					};
					OrgChartVerticalCanvas.Children.Add(managerBorder);
					Canvas.SetLeft(managerBorder, managerX);
					Canvas.SetTop(managerBorder, managerRowY);
					managerBorders.Add(managerBorder);

					// Draw right-angled arrow from director to manager
					double directorMidX = directorX + directorNodeWidth / 2;
					double directorBottomY = directorY + directorNodeHeight;
					double managerMidX = managerX + managerNodeWidth / 2;
					double managerTopY = managerRowY;

					// Right-angled: down from director, then horizontal to manager
					double verticalY = directorBottomY + 16;
					var polyline = new Polyline
					{
						Stroke = Brushes.Black,
						StrokeThickness = 1.5,
						StrokeDashArray = new DoubleCollection { 4, 2 },
						Points = new PointCollection
												{
													new Point(directorMidX, directorBottomY),
													new Point(directorMidX, verticalY),
													new Point(managerMidX, verticalY),
													new Point(managerMidX, managerTopY)
												}
					};
					OrgChartVerticalCanvas.Children.Add(polyline);
					DrawArrowHead(managerMidX, managerTopY, managerMidX, verticalY, OrgChartVerticalCanvas);

					// Draw employees vertically under manager, max 8 per column
					var employees = manager.Employees;
					if (employees != null && employees.Count > 0)
					{
						for (int col = 0; col < columns; col++)
						{
							for (int row = 0; row < batchSize; row++)
							{
								int idx = col * batchSize + row;
								if (idx >= employees.Count) break;
								var member = employees[idx];
								double employeeX = nextManagerX + col * (employeeNodeWidth + employeeSpacingX);
								double employeeY = managerRowY + managerNodeHeight + 40 + row * (employeeNodeHeight + employeeSpacingY);

								var employeeBorder = new Border
								{
									Width = employeeNodeWidth,
									Height = employeeNodeHeight,
									Background = member.Background ?? Brushes.White,
									BorderBrush = Brushes.Black,
									BorderThickness = new Thickness(1),
									CornerRadius = new CornerRadius(8),
									Child = new StackPanel
									{
										VerticalAlignment = VerticalAlignment.Center,
										HorizontalAlignment = HorizontalAlignment.Center,
										Children =
																{
																	new TextBlock
																	{
																		Text = member.Name,
																		FontWeight = FontWeights.SemiBold,
																		FontSize = 8,
																		Foreground = Brushes.Black,
																		TextAlignment = TextAlignment.Center,
																		HorizontalAlignment = HorizontalAlignment.Center
																	},
																	new TextBlock
																	{
																		Text = member.RoleTitle,
																		FontSize = 7,
																		Foreground = Brushes.Black,
																		TextAlignment = TextAlignment.Center,
																		HorizontalAlignment = HorizontalAlignment.Center
																	}
																}
									}
								};
								OrgChartVerticalCanvas.Children.Add(employeeBorder);
								Canvas.SetLeft(employeeBorder, employeeX);
								Canvas.SetTop(employeeBorder, employeeY);

								// Only draw arrow for the first employee in each column
								if (row == 0)
								{
									double managerBottomX = managerMidX;
									double managerBottomY = managerRowY + managerNodeHeight;
									double employeeTopX = employeeX + employeeNodeWidth / 2;
									double employeeTopY = employeeY;

									// Right-angled: down from manager, then horizontal to employee
									double verticalEmpY = managerBottomY + 14;
									var empPolyline = new Polyline
									{
										Stroke = Brushes.Black,
										StrokeThickness = 1,
										StrokeDashArray = new DoubleCollection { 4, 2 },
										Points = new PointCollection
																{
																	new Point(managerBottomX, managerBottomY),
																	new Point(managerBottomX, verticalEmpY),
																	new Point(employeeTopX, verticalEmpY),
																	new Point(employeeTopX, employeeTopY)
																}
									};
									OrgChartVerticalCanvas.Children.Add(empPolyline);
									DrawArrowHead(employeeTopX, employeeTopY, employeeTopX, verticalEmpY, OrgChartVerticalCanvas);
								}
							}
						}
					}
					nextManagerX += managerWidth + directorSpacingX;
				}
			}
		}
		*/
		private void DrawOrgChartHorizontalLayout()
		{
			OrgChartCanvas.Children.Clear();
			double startX = 60, startY = 50, nodeWidth = 150, nodeHeight = 70, hSpacing = 40, vSpacing = 30;
			foreach (var offShoreAdmTeam in _viewModel)
			{
				DrawOrgChart(offShoreAdmTeam, startX, ref startY, nodeWidth, nodeHeight, hSpacing, vSpacing);
				startY = GetNextPageStartY(startY);
			}
			
		}

		private void DrawOrgChart(
			OrgChartViewModel offShoreAdmTeam,
			double x,
			ref double y,
			double width,
			double height,
			double hSpacing,
			double vSpacing)
		{
			double currentY = y,
				managerNodeHeight = 22,
				managerNodeWidth = 165,
				offShoreAdmWidth = 190,
				directorNodeHeight = 22,
				directorNodeWidth = 200,
				managerStartX = x + 60,
				employeeStartX = x +260,
				employeeNodeWidth = 162,
				employeeNodeHeight = 32,
				employeeSpacingX = 8,
				employeeSpacingY = 8,
				offShoreAdmStartX = x + 550;
			int batchSize = 6;


			// Summary Node
			var admSummaryBorder = new Border
			{
				Width = employeeNodeWidth,
				Height = 90,
				Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffe7b3")),
				BorderBrush = Brushes.Black,
				BorderThickness = new Thickness(1),
				CornerRadius = new CornerRadius(8),
				Child = new StackPanel
				{
					VerticalAlignment = VerticalAlignment.Center,
					HorizontalAlignment = HorizontalAlignment.Center,
					Children =
									{
										new TextBlock
										{
											Text = $"On Seat - {offShoreAdmTeam.OnSeatCount}",
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										},
										new TextBlock
										{
											Text = $"(2U - {offShoreAdmTeam.TwoUCount}, 1U - {offShoreAdmTeam.OneUCount}, 0U - {offShoreAdmTeam.ZeroUCount}, HA - {offShoreAdmTeam.HireAheadCount})",
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										},
										new TextBlock
										{
											Text = $"GD - {offShoreAdmTeam.GDCount}",
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										},
										new TextBlock
										{
											Text = $"Offered - {offShoreAdmTeam.OfferedCount}",
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										},
										new TextBlock
										{
											Text = $"Open - {offShoreAdmTeam.OpenCount}",
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										},
										new TextBlock
										{
											Text = $"Total - {offShoreAdmTeam.TotalCount}",
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Left,
											HorizontalAlignment = HorizontalAlignment.Left
										}
									}
				}
			};
			OrgChartCanvas.Children.Add(admSummaryBorder);
			Canvas.SetLeft(admSummaryBorder, x + 1100);
			Canvas.SetTop(admSummaryBorder, currentY);


			// Offshore ADM node
			var offShoreAdmBorder = new Border
			{
				Width = offShoreAdmWidth,
				Height = employeeNodeHeight,
				Background = Brushes.Yellow,
				BorderBrush = Brushes.Black,
				BorderThickness = new Thickness(1),
				CornerRadius = new CornerRadius(8),
				Child = new StackPanel
				{
					VerticalAlignment = VerticalAlignment.Center,
					HorizontalAlignment = HorizontalAlignment.Center,
					Children =
									{
										new TextBlock
										{
											Text = offShoreAdmTeam.OffShoreAdm,
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Center,
											HorizontalAlignment = HorizontalAlignment.Center
										},
										new TextBlock
										{
											Text = "ADM",
											FontSize = 8,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Center,
											HorizontalAlignment = HorizontalAlignment.Center
										}
									}
				}
			};
			OrgChartCanvas.Children.Add(offShoreAdmBorder);
			Canvas.SetLeft(offShoreAdmBorder, offShoreAdmStartX);
			currentY += 15;
			Canvas.SetTop(offShoreAdmBorder, currentY);

			var offShoreAdmBottomMidX = offShoreAdmStartX + (offShoreAdmWidth / 2);
			var offShoreAdmBottomMidY = currentY + employeeNodeHeight;
			
			// change Y to start of first director container
			currentY += employeeNodeHeight + 50;

			// Line connecting offshore adm and first director container
			var offShoreAdmLine = new Polyline
			{
				Stroke = Brushes.Black,
				StrokeThickness = 1,
				StrokeDashArray = new DoubleCollection { 4, 2 },
				Points = new PointCollection
								{
									new Point(offShoreAdmBottomMidX, offShoreAdmBottomMidY),
									new Point(offShoreAdmBottomMidX, currentY)
								}
			};
			OrgChartCanvas.Children.Add(offShoreAdmLine);
			var firstContainerTouchPoint = offShoreAdmLine.Points[1];
			var offShoreAdmBottomMidPoint = offShoreAdmLine.Points[0];
			// DrawArrowHead expects (x, y) as arrow tip, (fromX, fromY) as direction
			DrawArrowHead(firstContainerTouchPoint.X, firstContainerTouchPoint.Y, offShoreAdmBottomMidPoint.X, offShoreAdmBottomMidPoint.Y, OrgChartCanvas);

			foreach (var director in offShoreAdmTeam.DirectorNodes)
			{
				int managerCount = director.Managers?.Count ?? 0;
				//var employeeRowsCount = director.Managers.Sum(s => s.Employees.Count <= batchSize ? 1 : (s.Employees.Count / batchSize) + (s.Employees.Count % batchSize == 0 ? 0 : 1));
				var employeeRowsCount = director.Managers.Sum(s => s.EmployeeRows);
				double directorRectHeight = directorNodeHeight + employeeRowsCount * (employeeNodeHeight + employeeSpacingY) + vSpacing;

				if (currentY + directorRectHeight + 15 > CeilToPdfPageHeight(currentY))
				{
					DrawLegend(OrgChartCanvas, x, currentY);
					currentY = GetNextPageStartY(currentY);
				}
				// Draw director container rectangle
				var directorContainer = new Border
				{
					Width = 1280,
					Height = directorRectHeight,
					Background = Brushes.Transparent,
					BorderBrush = Brushes.Gray,
					BorderThickness = new Thickness(2),
					CornerRadius = new CornerRadius(18)
				};
				OrgChartCanvas.Children.Add(directorContainer);
				Canvas.SetLeft(directorContainer, x);
				Canvas.SetTop(directorContainer, currentY);

				// Draw director node (top left corner)
				var directorNode = new Border
				{
					Width = directorNodeWidth,
					Height = directorNodeHeight,
					Background = director.Background ?? Brushes.LightYellow,
					BorderBrush = Brushes.Black,
					BorderThickness = new Thickness(2),
					CornerRadius = new CornerRadius(8),
					Child = new StackPanel
					{
						Margin = new Thickness(4, 0, 0, 0),
						VerticalAlignment = VerticalAlignment.Center,
						HorizontalAlignment = HorizontalAlignment.Center,
						Children =
						{
							new TextBlock
							{
								Text = $"{director.Name}, {director.RoleTitle}",
								FontWeight = FontWeights.Bold,
								FontSize = 9,
								Foreground = Brushes.Black,
								TextAlignment = TextAlignment.Center,
								VerticalAlignment = VerticalAlignment.Center,
								HorizontalAlignment = HorizontalAlignment.Center
							}
						}
					}
				};
				OrgChartCanvas.Children.Add(directorNode);
				Canvas.SetLeft(directorNode, x + 10);
				Canvas.SetTop(directorNode, currentY + 10);

				double managerStartY = currentY + directorNodeHeight + 20;
				foreach (var managerObj in director.Managers)
				{
					var manager = managerObj as ManagerNode ?? managerObj;
					var employees = manager.Employees;
					double nodeY = managerStartY;
					if (employees != null && employees.Count > 0)
					{
						for (int i = 0; i < employees.Count; i++)
						{
							int row = i / batchSize;
							int col = i % batchSize;
							double nodeX = employeeStartX + col * (employeeNodeWidth + employeeSpacingX);
							nodeY = managerStartY + row * (employeeNodeHeight + employeeSpacingY);

							var member = employees[i];
							var employeeBorder = new Border
							{
								Width = employeeNodeWidth,
								Height = employeeNodeHeight,
								Background = member.Background ?? Brushes.White,
								BorderBrush = Brushes.Black,
								BorderThickness = new Thickness(1),
								CornerRadius = new CornerRadius(8),
								Child = new StackPanel
								{
									VerticalAlignment = VerticalAlignment.Center,
									HorizontalAlignment = HorizontalAlignment.Center,
									Children =
									{
										new TextBlock
										{
											Text = member.Name,
											FontWeight = FontWeights.SemiBold,
											FontSize = 9,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Center,
											HorizontalAlignment = HorizontalAlignment.Center
										},
										new TextBlock
										{
											Text = member.RoleTitle,
											FontSize = 8,
											Foreground = Brushes.Black,
											TextAlignment = TextAlignment.Center,
											HorizontalAlignment = HorizontalAlignment.Center
										}
									}
								}
							};
							OrgChartCanvas.Children.Add(employeeBorder);
							Canvas.SetLeft(employeeBorder, nodeX);
							Canvas.SetTop(employeeBorder, nodeY - 5);
						}
					}
					if (!string.IsNullOrWhiteSpace(manager.Name))
					{
						var managerBorder = new Border
						{
							Width = managerNodeWidth, // Reduced width for smaller font
							Height = managerNodeHeight,
							Background = manager.Background ?? Brushes.LightGoldenrodYellow,
							BorderBrush = Brushes.Black,
							BorderThickness = new Thickness(1.5),
							CornerRadius = new CornerRadius(7),
							Child = new StackPanel
							{
								Margin = new Thickness(4, 0, 0, 0),
								VerticalAlignment = VerticalAlignment.Center,
								HorizontalAlignment = HorizontalAlignment.Center,
								Children =
								{
									new TextBlock
									{
										Text = $"{manager.Name}, {manager.RoleTitle}",
										FontWeight = FontWeights.SemiBold,
										FontSize = 9,
										Foreground = Brushes.Black,
										TextAlignment = TextAlignment.Center,
										VerticalAlignment = VerticalAlignment.Center,
										HorizontalAlignment = HorizontalAlignment.Center
									}
								}
							}
						};
						OrgChartCanvas.Children.Add(managerBorder);
						Canvas.SetLeft(managerBorder, managerStartX);
						Canvas.SetTop(managerBorder, managerStartY);
						// Draw dotted arrow from director to manager
						double directorLeftX = x + 10; // left side of director node
						double directorBottomY = currentY + 10 + directorNodeHeight;
						double managerLeftX = managerStartX;
						double managerRightX = managerLeftX + managerNodeWidth;
						double managerCenterY = managerStartY + managerNodeHeight / 2;
						for (int j = 0; j < manager.EmployeeRows; j++)
						{
							var managerLine = new Polyline
							{
								Stroke = Brushes.Black,
								StrokeThickness = 1,
								StrokeDashArray = new DoubleCollection { 4, 2 },
								Points = new PointCollection
								{
									new Point(managerRightX, managerCenterY),
									new Point(managerRightX + 10, managerCenterY),
									new Point(managerRightX + 10, managerCenterY + j * (employeeNodeHeight + employeeSpacingY)),
									new Point(employeeStartX, managerCenterY + j * (employeeNodeHeight + employeeSpacingY))
								}
							};
							OrgChartCanvas.Children.Add(managerLine);
							int lastIdx = managerLine.Points.Count - 1;
							if (lastIdx > 0)
							{
								var end = managerLine.Points[lastIdx];
								var prev = managerLine.Points[lastIdx - 1];
								// DrawArrowHead expects (x, y) as arrow tip, (fromX, fromY) as direction
								DrawArrowHead(end.X, end.Y, prev.X, prev.Y, OrgChartCanvas);
							}

						}
						// Offset from left corner
						double directorArrowStartX = directorLeftX + 8;
						// Midpoint for right angle
						double midX = directorArrowStartX + (managerLeftX - directorArrowStartX) / 2;
						var polyline = new Polyline
						{
							Stroke = Brushes.Black,
							StrokeThickness = 1.5,
							StrokeDashArray = new DoubleCollection { 4, 2 },
							Points = new PointCollection
							{
								new Point(midX, directorBottomY),
								new Point(midX, managerCenterY),
								new Point(managerLeftX, managerCenterY)
							}
						};
						OrgChartCanvas.Children.Add(polyline);
						DrawArrowHead(managerLeftX, managerCenterY, midX, managerCenterY, OrgChartCanvas);
					}
					managerStartY = nodeY + employeeNodeHeight + employeeSpacingY;
				}
				currentY += directorRectHeight + vSpacing;
			}
			DrawLegend(OrgChartCanvas, x, currentY);
			y = currentY;
			OrgChartCanvas.MinHeight = currentY + 50;
		}
		
		private void DrawArrowHead(double x, double y, double fromX, double fromY, Canvas canvas)
		{
			// Arrow size
			double arrowLength = 12;
			double arrowWidth = 8;

			// Calculate direction
			double dx = x - fromX;
			double dy = y - fromY;
			double length = Math.Sqrt(dx * dx + dy * dy);

			if (length == 0) return;

			// Unit vector
			double ux = dx / length;
			double uy = dy / length;

			// Base of the arrowhead
			double baseX = x - arrowLength * ux;
			double baseY = y - arrowLength * uy;

			// Perpendicular vector
			double perpX = -uy;
			double perpY = ux;

			// Points of the arrowhead triangle
			Point p1 = new Point(x, y);
			Point p2 = new Point(baseX + arrowWidth / 2 * perpX, baseY + arrowWidth / 2 * perpY);
			Point p3 = new Point(baseX - arrowWidth / 2 * perpX, baseY - arrowWidth / 2 * perpY);

			var arrowHead = new Polygon
			{
				Points = new PointCollection { p1, p2, p3 },
				Fill = Brushes.Black
			};
			canvas.Children.Add(arrowHead);
		}

		private void ExportPdfButton_Click(object sender, RoutedEventArgs e)
		{
			var dlg = new SaveFileDialog
			{
				Filter = "PDF Files (*.pdf)|*.pdf",
				Title = "Export Org Chart as PDF"
			};
			if (dlg.ShowDialog() == true)
			{
				ExportCanvasToPdf(((TabItem)OrgChartTabControl.SelectedValue).Header.ToString() == "Horizontal" ? OrgChartCanvas : OrgChartVerticalCanvas, dlg.FileName);
				MessageBox.Show("Exported to PDF successfully.");
			}
		}

		private static readonly (string Label, string ColorHex)[] LegendItems =
		{
			("2U", "#FFFF00"),
			("1U", "#FF69B4"),
			("0U", "#FFFFFF"),
			("Hire Ahead", "#DDEEFF"),
			("Offered", "#FFD700"),
			("GD", "#B7E1A1"),
			("Open", "#A9A9A9"),
			("Open Hire Ahead", "#E6F2FF")
		};

		private void DrawLegend(Canvas canvas, double startX, double startY)
		{
			double rectWidth = 18;
			double rectHeight = 14;
			double gapRectToText = 6;
			double gapBetweenItems = 18;
			double fontSize = 13;
			double x = startX;

			foreach (var (label, colorHex) in LegendItems)
			{
				// colored box
				var box = new Border
				{
					Width = rectWidth,
					Height = rectHeight,
					Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex)),
					BorderBrush = Brushes.Black,
					BorderThickness = new Thickness(1),
					CornerRadius = new CornerRadius(2)
				};
				canvas.Children.Add(box);
				Canvas.SetLeft(box, x);
				Canvas.SetTop(box, startY);

				// label
				var text = new TextBlock
				{
					Text = label,
					FontSize = fontSize,
					FontWeight = FontWeights.SemiBold,
					Foreground = Brushes.Black
				};
				canvas.Children.Add(text);
				Canvas.SetLeft(text, x + rectWidth + gapRectToText);
				Canvas.SetTop(text, startY - 2);

				// measure to advance x precisely
				double pixelsPerDip = VisualTreeHelper.GetDpi(canvas).PixelsPerDip;
				var ft = new FormattedText(
					label,
					CultureInfo.CurrentUICulture,
					FlowDirection.LeftToRight,
					new Typeface("Segoe UI"),
					fontSize,
					Brushes.Black,
					pixelsPerDip);

				x += rectWidth + gapRectToText + ft.WidthIncludingTrailingWhitespace + gapBetweenItems;
			}
		}
		private void ExportCanvasToPdf(Canvas canvas, string filePath)
		{
			// Use canvas width for PDF page width, A4 height for page height
			double canvasWidth = canvas.ActualWidth;
			const double A4Height = 842; // 11.69 inch * 72

			// Render the entire canvas to a high DPI bitmap/quality i.e. 300 DPI
			double targetDpi = 300.0;
			double dpiScale = targetDpi / 72.0; // PDF is 72 DPI
			// Target image width and Height
			int renderWidth = (int)(canvas.ActualWidth * targetDpi / 96.0);
			int renderHeight = (int)(canvas.ActualHeight * targetDpi / 96.0);

			var size = new Size(canvas.ActualWidth, canvas.ActualHeight);
			canvas.Measure(size);
			canvas.Arrange(new Rect(size));
			// Render entire canvas content as single image
			var rtb = new RenderTargetBitmap(renderWidth, renderHeight, targetDpi, targetDpi, PixelFormats.Pbgra32);
			rtb.Render(canvas);

			// Calculate how many vertical slices/pages are needed
			int sliceHeightPx = (int)(A4Height * dpiScale);
			int totalHeightPx = renderHeight;
			int totalWidthPx = renderWidth;
			int pageCount = (int)Math.Ceiling((double)totalHeightPx / sliceHeightPx);

			var pdf = new PdfDocument();

			for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
			{
				int srcY = pageIndex * sliceHeightPx;
				int srcHeight = Math.Min(sliceHeightPx, totalHeightPx - srcY);

				// Crop the bitmap for this page
				var croppedBitmap = new CroppedBitmap(
					rtb,
					new Int32Rect(
						0,
						srcY,
						totalWidthPx,
						srcHeight
					)
				);

				using (var croppedStream = new MemoryStream())
				{
					var croppedEncoder = new PngBitmapEncoder();
					croppedEncoder.Frames.Add(BitmapFrame.Create(croppedBitmap));
					croppedEncoder.Save(croppedStream);

					var page = pdf.AddPage();
					page.Width = canvasWidth * (targetDpi / 96.0) / dpiScale;
					page.Height = A4Height;

					using (var gfx = XGraphics.FromPdfPage(page))
					{
						var img = XImage.FromStream(() => new MemoryStream(croppedStream.ToArray()));

						// Calculate image height in PDF points
						double imgHeightPt = srcHeight / dpiScale;
						double imgWidthPt = totalWidthPx / dpiScale;

						// Align left
						double drawX = 0;
						double drawY = 0;

						gfx.DrawImage(img, drawX, drawY, imgWidthPt, imgHeightPt);
					}
				}
			}
			pdf.Save(filePath);
		}
		private static int CeilToPdfPageHeight(double value)
		{
			return (int)(Math.Ceiling(value / _a4PageHeight) * _a4PageHeight); ;
		}
		private static int GetNextPageStartY(double value)
		{
			return CeilToPdfPageHeight(value) + 15;
		}
	}
}