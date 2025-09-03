using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace UPS_OrgChart_WPF
{
	public class OrgNode : Node
	{
		public Brush Background { get; set; } = Brushes.White;
	}

	public class DirectorNode : OrgNode
	{
		public ObservableCollection<ManagerNode> Managers { get; set; } = new();
	}

	public class ManagerNode : OrgNode
	{
		public const int MaxEmployeesInRow = 6;
		public int EmployeeRows => (int)Math.Ceiling((double)Employees.Count / MaxEmployeesInRow);
		public ObservableCollection<OrgNode> Employees { get; set; } = new();
	}
	public class Node
	{
		public string ItcAdm { get; set; }
		public string OnsiteDirector { get; set; }
		public string OnsiteManager { get; set; }
		public string Name { get; set; }
		public string RoleTitle { get; set; }
		public string PayGrade { get; set; }
		public string Req { get; set; }
		public string Status { get; set; }
	}

	public class GradeMapConfiguration : ConfigurationElement
	{
		[ConfigurationProperty("friendlyName", IsRequired = true)]
		public string FriendlyName { get; set; }
		[ConfigurationProperty("name", IsRequired = true)]
		public string Name { get; set; }
		[ConfigurationProperty("colourCode", IsRequired = true)]
		public string ColourCode { get; set; }
	}
}
