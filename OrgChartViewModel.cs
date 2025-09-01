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
		public ObservableCollection<DirectorNode> DirectorNodes { get; set; } = new ObservableCollection<DirectorNode>();
		public OrgChartViewModel()
		{
			/*
			RootNodes = new ObservableCollection<OrgNode>();
			// Example based on your diagram
			var abc = new OrgNode { Name = "abc", Background = Brushes.LightBlue };
			var def = new OrgNode { Name = "def", Background = Brushes.LightGreen };
			var ghi = new OrgNode { Name = "ghi", Background = Brushes.LightYellow };

			var m = new OrgNode { Name = "m" };
			var n = new OrgNode { Name = "n" };
			var o = new OrgNode { Name = "o" };
			var p = new OrgNode { Name = "p" };
			var q = new OrgNode { Name = "q" };
			var r = new OrgNode { Name = "r" };

			ghi.Children.Add(m);
			ghi.Children.Add(n);
			ghi.Children.Add(o);
			ghi.Children.Add(p);
			ghi.Children.Add(q);
			ghi.Children.Add(r);

			abc.Children.Add(def);
			abc.Children.Add(ghi);

			var jkl = new OrgNode { Name = "jkl" };
			var s = new OrgNode { Name = "s" };
			var t = new OrgNode { Name = "t" };
			var u = new OrgNode { Name = "u" };

			jkl.Children.Add(s);
			jkl.Children.Add(t);
			jkl.Children.Add(u);

			def.Children.Add(jkl);

			RootNodes.Add(abc);
			//RootNodes.Add(def);
			*/
		}
	}
}
