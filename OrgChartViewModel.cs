using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace UPS_OrgChart_WPF
{
	public class OrgChartViewModel
	{
		public const string OnSeatRegular = "On seat - Regular";
		public const string OpenHireAhead = "Open - Hire Ahead";
		public const string OnSeatHireAhead = "On seat - Hire Ahead";
		public const string OfferAccepted = "Offer Accepted";
		public const string GD = "GD";
		public const string InterviewInProgress = "Interview in progress";

		public string OffShoreAdm { get; set; }

		public IEnumerable<OrgNode> AllEmployees => DirectorNodes.SelectMany(d => d.Managers).SelectMany(m => m.Employees);
		public int OnSeatCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OnSeatHireAhead, StringComparison.InvariantCultureIgnoreCase)
		|| c.Status.Equals(OrgChartViewModel.OnSeatRegular, StringComparison.InvariantCultureIgnoreCase));
		public int OpenCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.InterviewInProgress, StringComparison.InvariantCultureIgnoreCase)
		|| c.Status.Equals(OrgChartViewModel.OpenHireAhead, StringComparison.InvariantCultureIgnoreCase));
		public int OfferedCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OfferAccepted, StringComparison.InvariantCultureIgnoreCase));
		public int GDCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.GD, StringComparison.InvariantCultureIgnoreCase));
		public int TwoUCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OnSeatRegular, StringComparison.InvariantCultureIgnoreCase) && c.PayGrade.Equals("2U", StringComparison.InvariantCultureIgnoreCase));
		public int  OneUCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OnSeatRegular, StringComparison.InvariantCultureIgnoreCase) && c.PayGrade.Equals("1U", StringComparison.InvariantCultureIgnoreCase));
		public int ZeroUCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OnSeatRegular, StringComparison.InvariantCultureIgnoreCase) && c.PayGrade.Equals("0U", StringComparison.InvariantCultureIgnoreCase));
		public int HireAheadCount => AllEmployees.Count(c => c.Status.Equals(OrgChartViewModel.OnSeatHireAhead, StringComparison.InvariantCultureIgnoreCase));
		public int TotalCount => AllEmployees.Count();
		public ObservableCollection<DirectorNode> DirectorNodes { get; set; } = new ObservableCollection<DirectorNode>();
		public OrgChartViewModel()
		{
		}
	}
}
