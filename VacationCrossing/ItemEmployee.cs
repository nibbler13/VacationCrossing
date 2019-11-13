using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VacationCrossing {
	class ItemEmployee {
		public int ID { get; set; } = -1;
		public string Department { get; set; } = string.Empty;
		public string Name { get; set; } = string.Empty;
		public string Position { get; set; } = string.Empty;
		public string Type { get; set; } = string.Empty;
		public string SheetName { get; set; } = string.Empty;
		public List<Tuple<int, DateTime>> vacationPeriods { get; set; } = new List<Tuple<int, DateTime>>();
	}
}
