using System;
using System.Text.RegularExpressions;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace Varan.FrameMaker {
	public class MIF {
		protected MIFSettings settings;

		public MIF(MIFSettings initialSettings) {
			settings = initialSettings;
		}
		public void ConvertBookToPDF(string sourceBook, string targetBook) {
			//create command file for dzbatcher
			StreamWriter sw = new StreamWriter(settings.CommandFile, false);

			sw.WriteLine("Open " + sourceBook);
			sw.WriteLine("SaveAs -p " + sourceBook + " " + targetBook);
			sw.WriteLine("Close " + sourceBook);
			sw.Close();

			//run dzbatcher
			string dzbatcher = settings.DZBatcherEXE;
			Process p = Process.Start(dzbatcher, "-v " + settings.CommandFile);
			p.WaitForExit();
		}

		public void ConvertFMtoMIF(string sourceFolder, string targetFolder, ArrayList filenames) {
			//create command file for dzbatcher
			StreamWriter sw = new StreamWriter(settings.CommandFile, false);

			foreach(string file in filenames) {
				sw.WriteLine("Open " + sourceFolder + "\\" + file);
				sw.WriteLine("SaveAs -m " + sourceFolder + "\\" + file + " " + targetFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".mif");
				sw.WriteLine("Close " + sourceFolder + "\\" + file);
			}
			sw.Close();

			//run dzbatcher
			string dzbatcher = settings.DZBatcherEXE;
			Process p = Process.Start(dzbatcher, "-v " + settings.CommandFile);
			p.WaitForExit();
		}

		public void ConvertMIFtoFM(string sourceFolder, string targetFolder, ArrayList filenames) {
			StreamWriter sw = new StreamWriter(settings.CommandFile, false);

			foreach(string file in filenames) {
				sw.WriteLine("OpenTextFile -a " + sourceFolder + "\\" + file);
				sw.WriteLine("SaveAs -d " + sourceFolder + "\\" + file + " " + targetFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".fm");
				sw.WriteLine("Close " + sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".fm");
			}
			sw.Close();

			//run dzbatcher
			Process p = Process.Start(settings.DZBatcherEXE, "-v " + settings.CommandFile);
			p.WaitForExit();
		}

		public void AddMarker(StreamWriter sw, string markerType, string markerName, string markerText) {
			sw.WriteLine("<Marker");
			sw.WriteLine("<MType " + markerType + ">");
			sw.WriteLine("<MTypeName `" + markerName + "'>");
			sw.WriteLine("<MText `" + markerText + "'>");
			sw.WriteLine("> # end of Marker");
		}

		public void RemoveMarker() {
		}

		public void AddMarkersForCharacterTags(string sourceFolder, ArrayList filenames) {
			//add markers
			XmlNodeList items = settings.MIFSettingsXML.SelectNodes("/MIF_SETTINGS/MARKERS_FROM_CHARACTER_TAGS//ITEM");
			foreach(XmlNode item in items) {
				string markerType = item.Attributes["markerType"].Value;
				string markerName = item.Attributes["markerName"].Value;
				string characterTag = item.Attributes["characterTag"].Value;
				string sublevelParaTag = item.Attributes["sublevelParaTag"].Value;
				string boldParaTags = item.Attributes["boldParaTags"].Value;
				string normalPageCharTag = item.Attributes["normalPageCharTag"].Value;
				string boldPageCharTag = item.Attributes["boldPageCharTag"].Value;
				string customSort = item.Attributes["customSort"].Value;
				bool includeParaNum = Convert.ToBoolean(Convert.ToInt32(item.Attributes["includeParaNum"].Value));
				
				ArrayList XRefs = new ArrayList();
				XmlNodeList crossrefs = settings.MIFSettingsXML.SelectNodes("/MIF_SETTINGS/MARKERS_FROM_CHARACTER_TAGS/ITEM[@characterTag='" + characterTag + "']//CROSSREF");
				foreach(XmlNode xref in crossrefs) {
					XRefs.Add(xref.Attributes["keyword"].Value + ":" + xref.Attributes["see"].Value);
				}

				foreach(string file in filenames) {
					//all markers for the given character tag are removed and rebuilt, in order to update them
					RemoveMarkersFromOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerName);
					AddMarkersForCharacterTagsOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerType, markerName, characterTag, sublevelParaTag, boldParaTags, normalPageCharTag, boldPageCharTag, customSort, XRefs);
				}
			}
		}

		private void RemoveMarkersFromOneFile(string fileName, string tempFile, string markerName) {
			StreamReader sr = new StreamReader(fileName);
			StreamWriter sw = new StreamWriter(tempFile, false);

			string line, trimmedLine;
			bool doneWithHeader = false;

			//skip the header info
			line = sr.ReadLine();
			trimmedLine = line.Trim();
			while(!doneWithHeader) {
				if(trimmedLine == "> # end of Document") {
					doneWithHeader = true;
				} else {
					sw.WriteLine(line);
					line = sr.ReadLine();
					trimmedLine = line.Trim();
				}
			}
			sw.WriteLine(line);

			line = sr.ReadLine();
			while (line != null) {
				trimmedLine = line.Trim();

				// ADD CHECK HERE - don't remove ALL markers, just ones of this type.
				if(line != null && trimmedLine.StartsWith("<Marker")) {
					ArrayList markerInfo = new ArrayList();
					markerInfo.Add(line);

					//remove the existing Marker
					bool removeThisMarker = true;
					bool endOfMarker = false;
					while(!endOfMarker) {
						line = sr.ReadLine();
						trimmedLine = line.Trim();
						markerInfo.Add(line);
						if(trimmedLine.StartsWith("> # end of Marker")) {
							endOfMarker = true;
						}
						if(trimmedLine.StartsWith("<MTypeName `") && !trimmedLine.StartsWith("<MTypeName `" + markerName + "'>")) {
							removeThisMarker = false;
						}
					}
					if(!removeThisMarker) {
						//put it back!
						foreach(string s in markerInfo) {
							sw.WriteLine(s);
						}
					}
					line = sr.ReadLine();
				}else {
					sw.WriteLine(line);
					line = sr.ReadLine();
				}
			}
			sr.Close();
			sw.Close();
			File.Delete(fileName);
			File.Move(tempFile, fileName);
		}

		private void AddMarkersForCharacterTagsOneFile(string fileName, string tempFile, string markerType, 
						string markerName, string characterTag, string sublevelParaTag, string boldParaTags,
						string normalPageCharTag, string boldPageCharTag, string customSort, ArrayList XRefs) {
			StreamReader sr = new StreamReader(fileName);
			StreamWriter sw = new StreamWriter(tempFile, false);

			string line;
			string[] paraTags = sublevelParaTag.Split(";".ToCharArray());
			string[] boldTags = boldParaTags.Split(";".ToCharArray());

			string lastMajorParaTagText = "";
			string lastMajorParaTagNum = "";
            string tempLastMajorParaTagText = "";
			string lastParaTag = "";
			bool doneWithHeader = false;
            bool majorParaStringFound = false;
			bool lookingForParaTagText = false;
			
			//holds the current id of the table - if there is no table, set to zero
			int inTable = 0;

			while ((line = sr.ReadLine()) != null) {
				string trimmedLine = line.Trim();

                if (trimmedLine.Equals("<TblID 25>"))
                {
                    Debug.WriteLine("here");
                }


                if (trimmedLine.Contains("<Unique 1059122>"))
                {
                    Debug.WriteLine("here");
                }
            
                //skip the header info
				while(!doneWithHeader) {
					if(trimmedLine == "> # end of Document") {
						doneWithHeader = true;
					} else {
						sw.WriteLine(line);
						line = sr.ReadLine();
						trimmedLine = line.Trim();
					}
				}

				//done with header


                //table definition, set inTable (table id)
				if(trimmedLine.StartsWith("<TblID")) {
					int startIdx = trimmedLine.IndexOf(" ") + 1;
					int len = trimmedLine.LastIndexOf(">") - startIdx;
					inTable = Convert.ToInt32(trimmedLine.Substring(startIdx, len));
				}

				// if in a table definition, find the paragraph tag for items in that table
                if(inTable > 0) {
                    // return the title of the table, or empty string if it has no title or the title is not a paragraph style we're interested in
                    string tempTableTitle = GetTableTitle(fileName, inTable, paraTags);

                    // if the table title is empty, then use the last paragraph before the table entry (the "<ATbl" entry)
                    if (tempTableTitle.Length == 0)
                    {
                        string tempParaTagTextBeforeTable = CheckForParaTagOfTable(fileName, inTable, paraTags).Trim();
                        if (tempParaTagTextBeforeTable.Length > 0)
                        {
                            lastMajorParaTagText = tempParaTagTextBeforeTable;
                        }
                    }
                    else
                    {
                        lastMajorParaTagText = tempTableTitle.Trim();
                    }
					inTable = 0;
                // if not in a table definiton
                } else {
                         // if we find a new paragraph entry, see if it matches one we're intersted in
                        // if it does, set the lookingForParaTagText flag
                        if (trimmedLine.StartsWith("<PgfTag"))
                        {
                            int lastParaTagStartIdx = trimmedLine.IndexOf("`") + 1;
                            int lastParaTagLen = trimmedLine.LastIndexOf("'") - lastParaTagStartIdx;
                            lastParaTag = trimmedLine.Substring(lastParaTagStartIdx, lastParaTagLen);

                            lookingForParaTagText = false;

                            for (int i = 0; i < paraTags.Length; i++)
                            {
                                if (lastParaTag.Equals(paraTags[i]))
                                {
                                    lookingForParaTagText = true;
                                    majorParaStringFound = false;
                                    if(!lastParaTag.Equals("Title.Exercise"))
                                    {
                                        Debug.WriteLine("here");
                                    }
                                    break;
                                }
                            }
                        }

                        // if the lookingForParaTagText is set and we find a paragraph number, save the text for it
                        if (lookingForParaTagText && trimmedLine.StartsWith("<PgfNumString `"))
                        {
                            int lastParaNumStartIdx = trimmedLine.IndexOf("`") + 1;
                            int lastParaNumLen = trimmedLine.LastIndexOf("'") - lastParaNumStartIdx;
                            
                            // when we get a new paragraph number tag, save the current ones in case we need to back up to them (if the new one is blank)
                            tempLastMajorParaTagText = lastMajorParaTagText;

                            // assign the new ones
                            lastMajorParaTagNum = trimmedLine.Substring(lastParaNumStartIdx, lastParaNumLen);
                            lastMajorParaTagText = "";
                        }

                        // if the lookingForParaTagText is set and we find a paragraph string, append to the existing one (in case of multi-line)
                        if (lookingForParaTagText && trimmedLine.StartsWith("<String `"))
                        {
                            int startIdx = trimmedLine.IndexOf("`") + 1;
                            int len = trimmedLine.LastIndexOf("'") - startIdx;
                            if (trimmedLine.Substring(startIdx, len).Length > 0)
                            {
                                if (majorParaStringFound)
                                {
                                    // just append all subsequent <String tags
                                    lastMajorParaTagText += trimmedLine.Substring(startIdx, len);
                                }
                                else
                                {
                                    // prefix the para num for the first line (the first <String tag)
                                    lastMajorParaTagText = lastMajorParaTagNum + trimmedLine.Substring(startIdx, len);
                                }
                                majorParaStringFound = true;
                            } else {
                                // if the new ones were blank, then back up to the previous one
                                lastMajorParaTagText = tempLastMajorParaTagText;
                            }

                            if (lastMajorParaTagText.Equals("Specific Examples"))
                            {
                                Debug.WriteLine("here");
                            }
                        }

                        // once we reach the end of the paragraph entry, reset the lookingForParaTagText flag
                        if (lookingForParaTagText && trimmedLine.Equals("> # end of Para"))
                        {
                            lookingForParaTagText = false;
                            if (!majorParaStringFound)
                            {
                                lastMajorParaTagText = tempLastMajorParaTagText;
                            }
                        }
                    

                        // if we find a table entry, get the table title and see if we're interested in using it
                        // if we are, use it. Otherwise, use the last paragraph before the table entry that we were interested in
                        if (trimmedLine.StartsWith("<ATbl"))
                        {
                            int startIdx = trimmedLine.IndexOf(" ") + 1;
                            int len = trimmedLine.LastIndexOf(">") - startIdx;
                            int tableID = Convert.ToInt32(trimmedLine.Substring(startIdx, len));
                            string tempTableTitle = "";

                            tempTableTitle = GetTableTitle(fileName, tableID, paraTags).Trim();
                            if (tempTableTitle.Length > 0)
                            {
                                lastMajorParaTagText = tempTableTitle;
                                inTable = 0;
                            }
                        }
    			}

                //end of table definition, reset inTable (table id)
                if (trimmedLine.Equals("> # end of Tbl"))
                {
                    inTable = 0;
                }

				if(trimmedLine == "<FTag `" + characterTag + "'>") {
					//copy <FTag ...
					sw.WriteLine(line);

					bool doneWithCharTag = false;

					while(!doneWithCharTag) {
						line = sr.ReadLine();
						sw.WriteLine(line);

						trimmedLine = line.Trim();
						if(trimmedLine == "> # end of Font") {
							doneWithCharTag = true;
						}
					}

					bool doneWithMultiLine = false;
					string comboLine = "";
					ArrayList buffer = new ArrayList();
					string testStr;

					while(!doneWithMultiLine) {
						testStr = sr.ReadLine();
						buffer.Add(testStr);
						if(testStr.Trim().StartsWith("<String `")) {
							int startIdx = testStr.IndexOf("`") + 1;
							int len = testStr.LastIndexOf("'") - startIdx;
							comboLine += testStr.Substring(startIdx, len);
						}
						if(testStr.Trim().StartsWith("<FTag `'") || testStr.Trim().Equals("> # end of Para")) {
							doneWithMultiLine = true;
						}
					}

					string charTagText = ExtractCharTagText(comboLine);
                    if (charTagText.Equals("HT-7"))
                    {
                        Debug.WriteLine("here");
                    }
					string origCharTagText = "";
					string keyword = "";
					string see = "";

					int xrefIdx = -1;
					for(int i=0; i < XRefs.Count; i++) {
						if(XRefs[i].ToString().Trim().ToUpper().StartsWith(charTagText.Trim().ToUpper() + ":")) {
							xrefIdx = i;
							keyword = XRefs[xrefIdx].ToString().Substring(0, XRefs[xrefIdx].ToString().IndexOf(":"));
							see = XRefs[xrefIdx].ToString().Substring(XRefs[xrefIdx].ToString().IndexOf(":") + 1);
						}
					}

					if(xrefIdx >= 0) {
						origCharTagText = charTagText;
						charTagText = see;
					}

					string newParaText = "";
                    newParaText = lastMajorParaTagText.Trim().Replace(":", "-");

                    string charTagSortText = "";
                    charTagSortText = GetSortText(charTagText);
					
					if(newParaText.Length > 0) {
						charTagText += ":" + newParaText;
					}

					bool boldParaTagMatch = false;
					for(int i=0; i < boldTags.Length; i++) {
						if(boldTags[i].ToString() == lastParaTag) {
							boldParaTagMatch = true;
						}
					}
					
					if(boldParaTagMatch) {
						charTagText += "<" + boldPageCharTag + "\\>";
					}
					
					//check for custom sort
					if(customSort == "Acupoints" && xrefIdx < 0) {
						charTagText += charTagSortText;
					}

					if(customSort == "Acupoints" && xrefIdx >= 0) {
						charTagText = "<$nopage\\>" + keyword + ". <IdxSee\\>See " + see + GetSortText(keyword) + ";<$singlepage\\>" + charTagText + "</\\>" + charTagSortText;
					}

					AddMarker(sw, markerType, markerName, charTagText);
					foreach(string s in buffer) {
						sw.WriteLine(s);
					}
				} else {
					sw.WriteLine(line);
				}
			}
			sr.Close();
			sw.Close();
			File.Delete(fileName);
			File.Move(tempFile, fileName);
		}

		public string GetSortText(string charTagText) {
			bool isAcupoint = false;
			bool isCombo = false;

			string charTagSortText = "";
			if( charTagText.StartsWith("CV") ||
				charTagText.StartsWith("GV") ||
				charTagText.StartsWith("SI") ||
				charTagText.StartsWith("TH") ||
				charTagText.StartsWith("LI") ||
				charTagText.StartsWith("LU") ||
				charTagText.StartsWith("PC") ||
				charTagText.StartsWith("HT") ||
				charTagText.StartsWith("ST") ||
				charTagText.StartsWith("SP") ||
				charTagText.StartsWith("KI") ||
				charTagText.StartsWith("LV") ||
				charTagText.StartsWith("GB") ||
				charTagText.StartsWith("BL") ) {

				isAcupoint = true;
			}

			if(charTagText.IndexOf("/") >= 0) {
				isCombo = true;
			}
			
			if(isAcupoint && !isCombo) {
				try{
					string numText = "";
					numText = charTagText.Substring(charTagText.IndexOf("-")+1);
					int num = Convert.ToInt32(numText);
					string start = charTagText.Substring(0, charTagText.IndexOf("-")+1);
					string end = String.Format("{0:0#}", num);
					charTagSortText =  "[" + start + end + "]";
				} catch {
					charTagSortText = "[" + charTagText + "]";
				}
			}
					
			if(isAcupoint && isCombo) {
				try{
					string numText = "";
					int idxDash = charTagText.IndexOf("-");
					int idxSlash = charTagText.IndexOf("/");
					numText = charTagText.Substring(idxDash + 1, idxSlash - idxDash - 1);
					int num = Convert.ToInt32(numText);
					string start = charTagText.Substring(0, idxDash + 1);
					string end = charTagText.Substring(idxSlash);
					charTagSortText =  "[CMBO" + start + String.Format("{0:0#}", num) + end + "]";
				} catch {
					charTagSortText = "[CMBO" + charTagText + "]";
				}
			}

			if(!isAcupoint) {
				try{
					string numText = "";
					numText = charTagText.Substring(charTagText.IndexOf("-") + 1);
					int num = Convert.ToInt32(numText);
					string start = charTagText.Substring(0, charTagText.IndexOf("-") + 1);
					string end = String.Format("{0:0#}", num);
					charTagSortText =  "[EX" + start + end + "]";
				} catch {
					charTagSortText = "[EX" + charTagText + "]";
				}
			}

			return charTagSortText;
		}
		
		public string ExtractCharTagText(string line) {
			string charTagText = line.Trim();
			int startPos, len;
			startPos = "<String `".Length;
			len = charTagText.IndexOf("'>") - startPos;
			try {
				charTagText = charTagText.Substring(startPos, len);
			} catch(Exception exc) {
				string s = exc.Message;
			}

			//check for (R), (L), or (B) at end
			if(charTagText.Length > 3) {
				string suffix = charTagText.ToUpper().Substring(charTagText.Length - 3, 3);
				if( suffix == "(R)" || suffix == "(L)" || suffix == "(B)") {
					charTagText = charTagText.Substring(0, charTagText.Length - 3).Trim();
				}
			}

			return charTagText;
		}

		public ArrayList GetValidFMFiles(string path) {
			ArrayList fmFiles = new ArrayList();

			if(Directory.Exists(path) && path.IndexOf(" ") < 1) {
				string[] files = Directory.GetFiles(path, "*.fm");
				string excludePattern = ".backup.fm";
				
				for(int i=0; i<files.Length; i++) {
					if(!files[i].EndsWith(excludePattern)) {
						fmFiles.Add(Path.GetFileName(files[i]));
					}
				}
			}
			return fmFiles;
		}

		private string CheckForParaTagOfTable(string filename, int inTable, string[] paraTags) {
			File.Copy(filename, filename + ".tablescan", true);
			filename += ".tablescan";

            string paraTagText = "";
			string paraNum = "";
			StreamReader sr = new StreamReader(filename);

            bool inTextFlow = false;
			bool lookingForParaTagText = false;
			bool foundTable = false;
			bool inTblTitlePgf1 = false;
			
			//read down to the TextFlow
			string line = sr.ReadLine();
			while (!inTextFlow) {
				if(line.Trim().StartsWith("<TextFlow")) {
					inTextFlow = true;
				}
				line = sr.ReadLine();
			}

			string trimmedLine = "";
			while(line != null && !foundTable) {
				line = sr.ReadLine();
				trimmedLine = line.Trim();

                if(trimmedLine.StartsWith("<TblTitlePgf1")) {
					inTblTitlePgf1 = true;
				}

				if(trimmedLine.StartsWith("> # end of TblTitlePgf1")) {
					inTblTitlePgf1 = false;
				}

				if(!lookingForParaTagText && !inTblTitlePgf1) {
					for(int i=0; i < paraTags.Length; i++) {
						if(trimmedLine.StartsWith("<PgfTag `" + paraTags[i] + "'")) {
							lookingForParaTagText = true;
                            paraTagText = "";
							paraNum = "";
						}
					}
				}

				if(lookingForParaTagText && trimmedLine.Equals("> # end of Para")) {
						lookingForParaTagText = false;
				}

				if(lookingForParaTagText && trimmedLine.StartsWith("<String `")) {
					int startIdx = trimmedLine.IndexOf("`") + 1;
					int len = trimmedLine.LastIndexOf("'") - startIdx;
					paraTagText += trimmedLine.Substring(startIdx, len);
				}

				if(lookingForParaTagText && trimmedLine.StartsWith("<PgfNumString `")) {
					int startIdx = trimmedLine.IndexOf("`") + 1;
					int len = trimmedLine.LastIndexOf("'") - startIdx;
					paraNum = trimmedLine.Substring(startIdx, len);
				}

				if(trimmedLine.StartsWith("<ATbl " + inTable + ">")) {
					foundTable = true;
				}
				
			}
			sr.Close();
			
			
			paraTagText = paraNum + paraTagText;
			paraTagText = paraTagText.Trim();
            
            File.Delete(filename);
			return paraTagText;
		}

        private string GetTableTitle(string filename, int tableID, string[] paraTags)
        {
            string tableTitle = "";
            string tableTitleTag = "";
            string tableTitleParaNum = "";
            string line = "";
            bool foundTable = false;
            bool foundTableTitle = false;
            bool matchesWantedParaTag = false;

            File.Copy(filename, filename + ".tablescan", true);
            filename += ".tablescan";
            StreamReader sr = new StreamReader(filename);

            while ((line = sr.ReadLine()) != null)
            {
                if (line.Trim().StartsWith("<TblID " + tableID))
                {
                    foundTable = true;
                }
                if (foundTable && line.Trim().StartsWith("<TblTitleContent"))
                {
                    foundTableTitle = true;
                }
                if (foundTableTitle && line.Trim().StartsWith("<PgfTag"))
                {
                    tableTitleTag = line.Trim().Substring(9, line.Trim().Length - 11);
                }
                if (foundTable && foundTableTitle && line.Trim().StartsWith("<PgfNumString"))
                {
                    tableTitleParaNum = line.Trim().Substring(15, line.Trim().Length - 17);
                }
                if (foundTableTitle && line.Trim().StartsWith("<String"))
                {
                    tableTitle += line.Trim().Substring(9, line.Trim().Length - 11);
                }
                if (foundTableTitle && line.Trim().Equals("> # end of Para"))
                {
                    break;
                }
                if (foundTableTitle && line.Trim().StartsWith("> # end of TblTitleContent"))
                {
                    break;
                }
            }


            for (int i = 0; i < paraTags.Length; i++)
            {
                if (tableTitleTag.Equals(paraTags[i]))
                {
                    matchesWantedParaTag = true;
                }
            }

            sr.Close();
            File.Delete(filename);
            if (matchesWantedParaTag)
            {
                return tableTitleParaNum + tableTitle;
            }
            else
            {
                return "";
            }
        }

        public void AddMarkersForKeywords(string sourceFolder, ArrayList filenames) {
			//add markers
			XmlNode item = settings.MIFSettingsXML.SelectSingleNode("/IDX_KEYWORD_MARKERS/MARKER_SETTINGS");
			string markerType = item.Attributes["markerType"].Value;
			string markerName = item.Attributes["markerName"].Value;
			string normalPageCharTag = item.Attributes["normalPageCharTag"].Value;
			string boldPageCharTag = item.Attributes["boldPageCharTag"].Value;

			item = settings.MIFSettingsXML.SelectSingleNode("/IDX_KEYWORD_MARKERS/EXCEL_SETTINGS");
			string excelPath = item.Attributes["path"].Value;
			string excelSheetName = item.Attributes["sheetName"].Value;

			string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + 
				"Data Source=" + excelPath + ";" +
				"Extended Properties=\"Excel 8.0;HDR=YES\"";

			OleDbConnection conn = new OleDbConnection(connStr);

			string selectStr = "SELECT * FROM [" + excelSheetName + "$]";
			OleDbDataAdapter da = new OleDbDataAdapter(selectStr, conn);

			DataSet ds = new DataSet();
			da.SelectCommand.Connection.Open();
			da.Fill(ds, excelSheetName);
			da.SelectCommand.Connection.Close();			
			
			//ArrayList XRefs = new ArrayList();
			DataTable dtKeywords = ds.Tables[excelSheetName];

			foreach(string file in filenames) {
				RemoveMarkersFromOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerName);
				AddMarkersForKeywordOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerType, markerName, dtKeywords, boldPageCharTag);
			}
		}

		private void AddMarkersForKeywordOneFile(string fileName, string tempFile, string markerType, 
													string markerName, DataTable dtKeywords, string boldPageCharTag) {

			StreamReader sr = new StreamReader(fileName);
			StreamWriter sw = new StreamWriter(tempFile, false);
			string line;
			bool doneWithHeader = false;
			
			while ((line = sr.ReadLine()) != null) {
				string trimmedLine = line.Trim();
				
				//skip the header info
				while(!doneWithHeader) {
					if(trimmedLine == "> # end of Document") {
						doneWithHeader = true;
					} else {
						sw.WriteLine(line);
						line = sr.ReadLine();
						trimmedLine = line.Trim();
					}
				}
				if(trimmedLine.StartsWith("<String `")) {
					line = CheckForKeywords(line, dtKeywords, markerName, markerType, boldPageCharTag);
					string delimStr = "\n";
					char[] delimArray = delimStr.ToCharArray();
					string[] splitLines = line.Split(delimArray);
					for(int i=0; i < splitLines.Length; i++) {
						sw.WriteLine(splitLines[i]);
					}
				} else {
					sw.WriteLine(line);
				}
			}
			sr.Close();
			sw.Close();
			File.Delete(fileName);
			File.Move(tempFile, fileName);
		}

		private string CheckForKeywords(string line, DataTable dtKeywords, string markerName, string markerType, string boldPageCharTag) {
			string finalLine = line;
			
	
			foreach(DataRow dr in dtKeywords.Rows) {
				if(line.ToLower().IndexOf(dr["Keyword"].ToString().ToLower()) >= 0) {
					if(line.ToLower().IndexOf("leopard") >= 0) {
						Debug.WriteLine("found");
					}

					string markerText = dr["Keyword"].ToString();
					
					if(dr["SynonymFor"] != DBNull.Value) {
						if(dr["SeeEntry"] == DBNull.Value) {
							if(dr["ChildOf"] == DBNull.Value) {
								markerText = dr["SynonymFor"].ToString();
							} else {
								string childOf = dr["ChildOf"].ToString();
								markerText = childOf + ":" + dr["SynonymFor"].ToString();
							}
						} else {
							if(dr["ChildOf"] != DBNull.Value) {
								string childOf = dr["ChildOf"].ToString();
								markerText = "<$nopage\\>" + childOf + ":" + dr["Keyword"].ToString() + ". <Emphasis\\>See<Default Para Font\\> " + childOf + ", " + dr["SynonymFor"].ToString() + ";<$singlepage\\>" + childOf + ":" + dr["SynonymFor"].ToString();
							} else {
								markerText = "<$nopage\\>" + dr["Keyword"].ToString() + ". <Emphasis\\>See<Default Para Font\\> " + dr["SynonymFor"].ToString() + ";<$singlepage\\>" + dr["SynonymFor"].ToString();
							}
						}
					}

					if(dr["ChildOf"] != DBNull.Value && dr["SynonymFor"] == DBNull.Value) {
						markerText = dr["ChildOf"].ToString() + ":" + markerText;
					}

					if(dr["NoPage"] != DBNull.Value && dr["SynonymFor"] == DBNull.Value) {
						markerText = "<$nopage\\>" + markerText;
					}
					
					string replaceStr = "'><Marker\n<MType " + markerType + "><MTypeName `" + markerName + "'>\n<MText `" + markerText + "'>\n> # end of Marker\n<String `" + dr["Keyword"].ToString();
					finalLine = Regex.Replace(line, @"\b" + dr["Keyword"].ToString() + @"\b", replaceStr, RegexOptions.IgnoreCase);
				}
			}
			return finalLine;
		}

		public void AddMarkersForParagraphTags(string sourceFolder, ArrayList filenames) {
			//add markers
			XmlNodeList items = settings.MIFSettingsXML.SelectNodes("/MARKERS_FROM_PARAGRAPH_TAGS//ITEM");
			foreach(XmlNode item in items) {
				string markerType = item.Attributes["markerType"].Value;
				string markerName = item.Attributes["markerName"].Value;
				string paragraphTag = item.Attributes["paragraphTag"].Value;
				string sublevelParaTag = item.Attributes["sublevelParaTag"].Value;
				string boldParaTags = item.Attributes["boldParaTags"].Value;
				string normalPageCharTag = item.Attributes["normalPageCharTag"].Value;
				string boldPageCharTag = item.Attributes["boldPageCharTag"].Value;
				string customSort = item.Attributes["customSort"].Value;
				bool includeParaNum = Convert.ToBoolean(Convert.ToInt32(item.Attributes["includeParaNum"].Value));
				
				ArrayList XRefs = new ArrayList();
				XmlNodeList crossrefs = settings.MIFSettingsXML.SelectNodes("/MARKERS_FROM_PARAGRAPH_TAGS/ITEM[@paragraphTag='" + paragraphTag + "']//CROSSREF");
				foreach(XmlNode xref in crossrefs) {
					XRefs.Add(xref.Attributes["keyword"].Value + ":" + xref.Attributes["see"].Value);
				}

				foreach(string file in filenames) {
					//all markers for the given character tag are removed and rebuilt, in order to update them
					RemoveMarkersFromOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerName);
					AddMarkersForParagraphTagsOneFile(sourceFolder + "\\" + file, sourceFolder + "\\" + Path.GetFileNameWithoutExtension(file) + ".tmp", markerType, markerName, paragraphTag, sublevelParaTag, boldParaTags, normalPageCharTag, boldPageCharTag, customSort, includeParaNum, XRefs);
				}
			}
		}
		
		private void AddMarkersForParagraphTagsOneFile(string fileName, string tempFile, string markerType, 
			string markerName, string paragraphTag, string sublevelParaTag, string boldParaTags,
			string normalPageCharTag, string boldPageCharTag, string customSort, bool includeParaNum,
			ArrayList XRefs) {
			StreamReader sr = new StreamReader(fileName);
			StreamWriter sw = new StreamWriter(tempFile, false);

			string line;
			string[] paraTags = sublevelParaTag.Split(";".ToCharArray());
			string[] boldTags = boldParaTags.Split(";".ToCharArray());

			string lastMajorParaTagText = "";
			string lastMajorParaTagNum = "";
			string lastParaTag = "";
			string currentParaTag = "";
			bool doneWithHeader = false;
			bool lookingForParaTagText = false;
			bool inTblTitlePgf1 = false;
			
			//holds the current id of the table - if there is no table, set to zero
			int inTable = 0;

			while ((line = sr.ReadLine()) != null) {
				string trimmedLine = line.Trim();
				
				//skip the header info
				while(!doneWithHeader) {
					if(trimmedLine == "> # end of Document") {
						doneWithHeader = true;
					} else {
						sw.WriteLine(line);
						line = sr.ReadLine();
						trimmedLine = line.Trim();
					}
				}

				if(trimmedLine.StartsWith("<TblTitlePgf1")) {
					inTblTitlePgf1 = true;
				}

				if(trimmedLine.StartsWith("> # end of TblTitlePgf1")) {
					inTblTitlePgf1 = false;
				}

				if(trimmedLine.StartsWith("<PgfTag `")) {
					int startIdx = trimmedLine.IndexOf("`") + 1;
					int len = trimmedLine.LastIndexOf("'") - startIdx;
					lastParaTag = currentParaTag;
					currentParaTag = trimmedLine.Substring(startIdx, len);					
				}

				if(trimmedLine.StartsWith("<TblID")) {
					int startIdx = trimmedLine.IndexOf(" ") + 1;
					int len = trimmedLine.LastIndexOf(">") - startIdx;
					inTable = Convert.ToInt32(trimmedLine.Substring(startIdx, len));
				}

				if(trimmedLine.StartsWith(" > # end of Tbl")) {
					inTable = 0;
				}

				if(inTable > 0) {
					lastMajorParaTagText = CheckForParaTagOfTable(fileName, inTable, paraTags);
					inTable = 0;
				} else {

					if(!lookingForParaTagText && !inTblTitlePgf1) {
						for(int i=0; i < paraTags.Length; i++) {
							if(trimmedLine.StartsWith("<PgfTag `" + paraTags[i] + "'")) {
								lookingForParaTagText = true;
								lastMajorParaTagNum = "";
							}
						}
					}

					if(lookingForParaTagText && trimmedLine.Equals("> # end of Para")) {
						lookingForParaTagText = false;
					}

					if(lookingForParaTagText && trimmedLine.StartsWith("<String `")) {
						int startIdx = trimmedLine.IndexOf("`") + 1;
						int len = trimmedLine.LastIndexOf("'") - startIdx;
						lastMajorParaTagText = trimmedLine.Substring(startIdx, len);
					}

					if(lookingForParaTagText && trimmedLine.StartsWith("<PgfNumString `")) {
						int startIdx = trimmedLine.IndexOf("`") + 1;
						int len = trimmedLine.LastIndexOf("'") - startIdx;
						lastMajorParaTagNum = trimmedLine.Substring(startIdx, len);
					}
				}

				//*****
				string paragraphTagText = "";
				if(trimmedLine == "<PgfTag `" + paragraphTag + "'>") {
					//copy <PgfTag ...
					sw.WriteLine(line);

					bool doneWithParagraphTag = false;
					bool readingParaLineString = false;
					bool writeToBuffer = false;
					ArrayList buffer = new ArrayList();

					while(!doneWithParagraphTag) {
						line = sr.ReadLine();
						trimmedLine = line.Trim();
						
						if(trimmedLine.StartsWith("<String `")) {
							writeToBuffer = true;
						}

						if(writeToBuffer) {
							buffer.Add(line);
						} else {
							sw.WriteLine(line);
						}
						
						
						if(readingParaLineString) {
							if(trimmedLine.StartsWith("<String `")) {
//								if(trimmedLine.StartsWith("<String `This is a test")) {
//									int a = 1;
//								}
								int startIdx = trimmedLine.IndexOf("`") + 1;
								int len = trimmedLine.LastIndexOf("'") - startIdx;
								paragraphTagText += trimmedLine.Substring(startIdx, len);
							}
						}

						if(trimmedLine.StartsWith("<ParaLine")) {
							readingParaLineString = true;
						} else {
							readingParaLineString = false;
						}

						//*****
						
						if(trimmedLine.Equals("> # end of Para")) {
							doneWithParagraphTag = true;
						}
					}

					string origCharTagText = "";
					string keyword = "";
					string see = "";

					int xrefIdx = -1;
					for(int i=0; i<XRefs.Count; i++) {
						if(XRefs[i].ToString().Trim().ToUpper().StartsWith(paragraphTagText.Trim().ToUpper() + ":")) {
							xrefIdx = i;
							keyword = XRefs[xrefIdx].ToString().Substring(0, XRefs[xrefIdx].ToString().IndexOf(":"));
							see = XRefs[xrefIdx].ToString().Substring(XRefs[xrefIdx].ToString().IndexOf(":") + 1);
						}
					}

					if(xrefIdx >= 0) {
						origCharTagText = paragraphTagText;
						paragraphTagText = see;
					}

					string newParaText = "";
					if(includeParaNum) {
						newParaText = lastMajorParaTagNum + lastMajorParaTagText;
						newParaText = newParaText.Trim();
					} else {
						newParaText = lastMajorParaTagText.Trim();
					}

					//string paraTagSortText = GetSortText(paragraphTagText);
					
					if(newParaText.Length > 0) {
						paragraphTagText += ":" + newParaText;
					}

					bool boldParaTagMatch = false;
					for(int i=0; i < boldTags.Length; i++) {
						if(boldTags[i].ToString() == lastParaTag) {
							boldParaTagMatch = true;
						}
					}
					
					if(boldParaTagMatch) {
						paragraphTagText += "<" + boldPageCharTag + "\\>";
					}
					
					//check for custom sort
					AddMarker(sw, markerType, markerName, paragraphTagText);
					foreach(string s in buffer) {
						sw.WriteLine(s);
					}
				} else {
					sw.WriteLine(line);
				}
			}
			sr.Close();
			sw.Close();
			File.Delete(fileName);
			File.Move(tempFile, fileName);
		}
	}
}