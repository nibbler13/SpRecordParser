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
		public int callsAccidentialShort { get; set; }
		public int callsAccidentialWrongValues { get; set; }
		public int callsDialed { get; set; }

		public int ringUp1tryRegulationObserved { get; set; }
		public int ringUp1tryRegulationNotObserved { get; set; }
		public int ringUp2tryRegulationObserved { get; set; }
		public int ringUp2tryRegulationNotObserved { get; set; }
		public int ringUp3tryRegulationObserved { get; set; }
		public int ringUp3tryRegulationNotObserved { get; set; }
		public int ringUp3MoreTryRegulationObserved { get; set; }
		public int ringUp3MoreTryRegulationNotObserved { get; set; }
		public int ringUpByPatientRegulationObserved { get; set; }
		public int ringUpByPatientRegulationNotObserved { get; set; }
		public int ringUpNotRegulationObserved { get; set; }
		public int ringUpNotRegulationNotObserved { get; set; }
		public int ringUpDidNotTried { get; set; }
		
		public TimeSpan timeAccepted { get; set; }
		public TimeSpan timeMissed { get; set; }
		public TimeSpan timeAccidential { get; set; }
		public TimeSpan timeDialed { get; set; }

		public SpRecordFileInformation() {

		}
	}
}
