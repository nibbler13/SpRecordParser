using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpRecordParser {
	class SpRecordFileInformation {
		public int callsAccepted { get; set; }
		public int callsMissed { get; set; }
		public int callsDialed { get; set; }
		public int callsDialedBackTotal { get; set; }
		public int callsDialedBackSuccessed { get; set; }
		public int callsDialedBackSuccessedInTime { get; set; }
		public int callsAccidential { get; set; }
		public int callsTotal { get; set; }
		public string timeTotal { get; set; }
		public TimeSpan timeAccepted { get; set; }
		public TimeSpan timeMissed { get; set; }
		public TimeSpan timeDialed { get; set; }
		public string workstationName { get; set; }
		public string creationDate { get; set; }
		public string accountingPeriod { get; set; }
		public List<List<string>> fileContent { get; set; }

		public SpRecordFileInformation() {

		}
	}
}
