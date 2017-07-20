using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpRecordParser {
	class SpRecordFileInformation {
		public List<List<string>> fileContent { get; set; }

		//Заполняется значениеми из файла
		public int CallsTotal { get; set; }

		public string WorkstationName { get; set; }
		public string CreationDate { get; set; }
		public string AccountingPeriod { get; set; }

		public TimeSpan TimeTotal { get; set; }

		//Рассчитывается автоматически
		public int CallsAccepted { get; set; }
		public int CallsMissed { get; set; }
		public int CallsMissedAccidentialShort { get; set; }
		public int CallsMissedAccidentialWrongValues { get; set; }
		public int CallsMissedInternal { get; set; }
		public int CallsDialed { get; set; }

		public int RingUp1tryRegulationObserved { get; set; }
		public int RingUp1tryRegulationNotObserved { get; set; }
		public int RingUp2tryRegulationObserved { get; set; }
		public int RingUp2tryRegulationNotObserved { get; set; }
		public int RingUp3tryRegulationObserved { get; set; }
		public int RingUp3tryRegulationNotObserved { get; set; }
		public int RingUp3MoreTryRegulationObserved { get; set; }
		public int RingUp3MoreTryRegulationNotObserved { get; set; }
		public int RingUpByPatientRegulationObserved { get; set; }
		public int RingUpByPatientRegulationNotObserved { get; set; }
		public int RingUpNotRegulationObserved { get; set; }
		public int RingUpNotRegulationNotObserved { get; set; }
		public int RingUpDidNotTried { get; set; }
		
		public TimeSpan TimeAccepted { get; set; }
		public TimeSpan TimeMissed { get; set; }
		public TimeSpan TimeAccidential { get; set; }
		public TimeSpan TimeDialed { get; set; }

		public SpRecordFileInformation() {

		}
	}
}
