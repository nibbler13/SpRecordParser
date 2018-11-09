using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpRecordParser {
	class SpRecordFileInformation {
		public List<List<string>> FileContent { get; set; }

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

		public Dictionary<DateTime, DayInfo> DaysInfo { get; set; }

		public SpRecordFileInformation() {
			DaysInfo = new Dictionary<DateTime, DayInfo>();
		}

		public class DayInfo {
			public enum MissedCallType {
				ConditionalyLostSelf5sec,
				ConditionalyLostSelf10sec,
				ConditionalyLostSelf15sec,
				ConditionalyLostSelf20sec,
				ConditionalyLostSelf25sec,
				ConditionalyLostSelf30sec,
				ConditionalyLostRedirected5sec,
				ConditionalyLostRedirected10sec,
				ConditionalyLostRedirected15sec,
				ConditionalyLostRedirected20sec,
				ConditionalyLostRedirected25sec,
				ConditionalyLostRedirected30sec
			}

			public int TotalIncoming { get; set; }
			public int TotalRedirected { get; set; }

			private readonly Dictionary<MissedCallType, int> missedCalls = new Dictionary<MissedCallType, int>() {
				{MissedCallType.ConditionalyLostSelf5sec, 0 },
				{MissedCallType.ConditionalyLostSelf10sec, 0 },
				{MissedCallType.ConditionalyLostSelf15sec, 0 },
				{MissedCallType.ConditionalyLostSelf20sec, 0 },
				{MissedCallType.ConditionalyLostSelf25sec, 0 },
				{MissedCallType.ConditionalyLostSelf30sec, 0 },
				{MissedCallType.ConditionalyLostRedirected5sec, 0 },
				{MissedCallType.ConditionalyLostRedirected10sec, 0 },
				{MissedCallType.ConditionalyLostRedirected15sec, 0 },
				{MissedCallType.ConditionalyLostRedirected20sec, 0 },
				{MissedCallType.ConditionalyLostRedirected25sec, 0 },
				{MissedCallType.ConditionalyLostRedirected30sec, 0 }
			};

			public int TotalConditionalyLost {
				get {
					return GetMissedCallCount(MissedCallType.ConditionalyLostSelf15sec) + 
						GetMissedCallCount(MissedCallType.ConditionalyLostRedirected20sec);
				}
			}

			public double TotalConditionalyLostPercent {
				get {
					return (double)TotalConditionalyLost / (double)TotalIncoming;
				}
			}

			public void IncrementMissedCallCount(MissedCallType missedCallType) {
				missedCalls[missedCallType]++;
			}

			public int GetMissedCallCount(MissedCallType missedCallType) {
				switch (missedCallType) {
					case MissedCallType.ConditionalyLostSelf5sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf5sec];
					case MissedCallType.ConditionalyLostSelf10sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf5sec] +
							missedCalls[MissedCallType.ConditionalyLostSelf10sec];
					case MissedCallType.ConditionalyLostSelf15sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf15sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostSelf10sec);
					case MissedCallType.ConditionalyLostSelf20sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf20sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostSelf15sec);
					case MissedCallType.ConditionalyLostSelf25sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf25sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostSelf20sec);
					case MissedCallType.ConditionalyLostSelf30sec:
						return missedCalls[MissedCallType.ConditionalyLostSelf30sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostSelf25sec);
					case MissedCallType.ConditionalyLostRedirected5sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected5sec];
					case MissedCallType.ConditionalyLostRedirected10sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected5sec] +
							missedCalls[MissedCallType.ConditionalyLostRedirected10sec];
					case MissedCallType.ConditionalyLostRedirected15sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected15sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostRedirected10sec);
					case MissedCallType.ConditionalyLostRedirected20sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected20sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostRedirected15sec);
					case MissedCallType.ConditionalyLostRedirected25sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected25sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostRedirected20sec);
					case MissedCallType.ConditionalyLostRedirected30sec:
						return missedCalls[MissedCallType.ConditionalyLostRedirected30sec] +
							GetMissedCallCount(MissedCallType.ConditionalyLostRedirected25sec);
					default:
						return 0;
				}
			}
		}
	}
}
