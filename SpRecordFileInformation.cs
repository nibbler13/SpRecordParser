using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpRecordParser {
	class SpRecordFileInformation {
		public List<List<string>> fileContent { get; set; }

		//Заполняется значениеми из файла
		public int callsTotal { get; set; }

		public string workstationName { get; set; }
		public string creationDate { get; set; }
		public string accountingPeriod { get; set; }

		public TimeSpan timeTotal { get; set; }

		//Рассчитывается автоматически
		public int callsAccepted { get; set; }
		public int callsMissed { get; set; }
		public int callsAccidential { get; set; }
		public int callsDialed { get; set; }

		public int callsBackByRegistry { get; set; }
		public int callsBackByPatient { get; set; }
		public int callsBackNot { get; set; }

		public int missedCallsRegulationObserved { get; set; }
		public int missedCallsRegulationNotObserved { get; set; }

		public TimeSpan timeAccepted { get; set; }
		public TimeSpan timeMissed { get; set; }
		public TimeSpan timeAccidential { get; set; }
		public TimeSpan timeDialed { get; set; }

		public SpRecordFileInformation() {

		}
	}
}
