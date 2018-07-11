using System;
using System.Xml;

namespace Varan.FrameMaker {
	public class MIFSettings {
		string tempFolder;
		string commandFile;
		string dzbatcherEXE;
		XmlDocument mifSettingsXml;

		public string TempFolder  {
			get{return tempFolder;}
			set{tempFolder = value;}
		}
		
		public string CommandFile  {
			get{return commandFile;}
			set{commandFile = value;}
		}

		public string DZBatcherEXE  {
			get{return dzbatcherEXE;}
			set{dzbatcherEXE = value;}
		}

		public XmlDocument MIFSettingsXML  {
			get {return mifSettingsXml;}
			set {mifSettingsXml = value;}
		}

		public MIFSettings() {
		}
	}
}
