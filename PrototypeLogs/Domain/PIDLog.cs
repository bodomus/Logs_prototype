using System;
using System.Reflection;
using Medoc;

namespace Pathway.WPF.Models
{
	/// <summary>
	/// PidLog Model
	/// </summary>
	public class PIDLog
	{
		public string TimeStamp { get; set; }
		public string P { get; set; }
		public string I { get; set; }
		public string D { get; set; }
		public string Error { get; set; }
		public string SetPoint { get; set; }
		public string OldSetPoint { get; set; }
		public string Temperature1 { get; set; }
		public string Temperature2 { get; set; }
		public string DAC { get; set; }
		public string RealTemperature1 { get; set; }
		public string RealTemperature2 { get; set; }
		public string WaterTemp { get; set; }
		public string PCB { get; set; }
		public string Heatsink1Temp { get; set; }
		public string Heatsink2Temp { get; set; }
		public string TEC { get; set; }
        public string SensorMismatch { get; set; }
        

        public PIDLog()
		{ }

		public PIDLog(string timeStamp, string p, string i, string d,
			string error, string setPoint, string oldSetPoint, string temperature1, string temperature2,
			string dac, string realTemperature1, string realTemperature2, string waterTemp, string pcb, string heatsink1Temp, string heatsink2Temp, string tec)
		{
			TimeStamp = timeStamp;
			P = p;
			I = i;
			D = d;
			Error = error;
			SetPoint = setPoint;
			OldSetPoint = oldSetPoint;
			Temperature1 = temperature1;
			Temperature2 = temperature2;
			DAC = dac;
			RealTemperature1 = realTemperature1;
			RealTemperature2 = realTemperature2;
			PCB = pcb;
			WaterTemp = waterTemp;
			Heatsink1Temp = heatsink1Temp;
			Heatsink2Temp = heatsink2Temp;
			TEC = tec;
		}

		/// <summary>
		/// For csv file saving headers
		/// </summary>
		public static string TraceObjectProperties(Object obj, DeviceType deviceType)
		{
			string result = string.Empty;
			foreach(PropertyInfo pi in obj.GetType().GetProperties())
			{
				switch(deviceType)
				{
					case DeviceType.TCU:
						if(VerificationPropertiesTCU(pi))
							result += String.Format("{0},", pi.Name);
						break;
					case DeviceType.CTSA:
						if(VerificationPropertiesCTSA(pi))
							result += String.Format("{0},", pi.Name);
						break;
					case DeviceType.TSA2:
						if(VerificationPropertiesTSA2(pi))
							result += String.Format("{0},", pi.Name);
						break;
					default:
						result += String.Format("{0},", pi.Name);
						break;
				}

			}
			return result;
		}

		/// <summary>
		/// List of values for CSV saving
		/// </summary>
		public string ToString(Object obj, DeviceType deviceType)
		{
			string result = String.Empty;
			foreach(PropertyInfo pi in obj.GetType().GetProperties())
			{
				switch(deviceType)
				{
					case DeviceType.TCU:
						if(VerificationPropertiesTCU(pi))
							result += String.Format("{0},", pi.GetValue(obj, null));
						break;
					case DeviceType.CTSA:
						if(VerificationPropertiesCTSA(pi))
							result += String.Format("{0},", pi.GetValue(obj, null));
						break;
					case DeviceType.TSA2:
						if(VerificationPropertiesTSA2(pi))
							result += String.Format("{0},", pi.GetValue(obj, null));
						break;
					default:
						result += String.Format("{0},", pi.GetValue(obj, null));
						break;
				}
			}
			return result;
		}

		private static bool VerificationPropertiesTCU(PropertyInfo pi)
		{
			return pi.Name != "Heatsink1Temp" && pi.Name != "Heatsink2Temp";
		}

		private static bool VerificationPropertiesTSA2(PropertyInfo pi)
		{
			return pi.Name != "OldSetPoint" && 
				pi.Name != "RealTemperature1" && 
				pi.Name != "RealTemperature2" && 
				pi.Name != "Heatsink1Temp" && 
				pi.Name != "Heatsink2Temp" &&
				pi.Name != "TEC" && 
				pi.Name != "WaterTemp";
		}

		private static bool VerificationPropertiesCTSA(PropertyInfo pi)
		{
			return pi.Name != "OldSetPoint" &&
				pi.Name != "RealTemperature1" &&
				pi.Name != "RealTemperature2" &&
				pi.Name != "PCB" &&
				pi.Name != "TEC" && 
				pi.Name != "WaterTemp";
		}
	}
}