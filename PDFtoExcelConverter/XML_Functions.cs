﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace PDFtoExcelConverter
{
    class XML_Functions
    {

        //public static string xml_name = AppDomain.CurrentDomain.BaseDirectory + @"networksettings.xml";
        public static void Create_Kabel_XML_File(string xml_name, string[] bmk_names, int nocopies)
        {





            try
            {
                if (File.Exists(xml_name))
                {
                    File.Delete(xml_name);
                }
                if (!File.Exists(xml_name))
                {

                    System.Xml.XmlWriterSettings settings = new XmlWriterSettings();


                    settings.Indent = true;
                    settings.IndentChars = "\t";


                    using (XmlTextWriter writer = new XmlTextWriter(xml_name, Encoding.UTF8))//
                    {
                        writer.QuoteChar = '\"';

                        writer.WriteStartDocument();
                        writer.Formatting = Formatting.Indented;

                        writer.WriteStartElement("LabelStrip");

                        //Metainfo
                        writer.WriteStartElement("MetaInfo");
                        writer.WriteElementString("Name", "Kabelbeschriftungen");
                        writer.WriteElementString("Description", "Lietremarkierer 23 x 4 mm");
                        writer.WriteElementString("CreationTime", "2014-03-26T13:25:20");
                        writer.WriteElementString("ModificationTime", "2014-03-26T13:25:20");
                        writer.WriteElementString("PrintTime", "2017-02-15T07:32:54");
                        writer.WriteEndElement();

                        writer.WriteStartElement("StripBlocks");

                        writer.WriteElementString("Distance", "0");

                        //For Loop
                        for (int i = 0; i < bmk_names.Length - 1; i++)
                        {
                            if (!string.IsNullOrEmpty(bmk_names[i]))
                            {
                                for (int k = 0; k < nocopies; k++) 
                                {
                                    writer.WriteStartElement("StripBlock");
                                    writer.WriteStartElement("StripRows");
                                    writer.WriteStartElement("StripRow");
                                    writer.WriteElementString("Height", "4000");
                                    writer.WriteElementString("TopOffset", "1750");
                                    writer.WriteElementString("BottomOffset", "1750");
                                    writer.WriteStartElement("StripCells");
                                    writer.WriteStartElement("StripCell");
                                    writer.WriteElementString("Width", "23000");
                                    writer.WriteStartElement("Content");
                                    writer.WriteElementString("Type", "Text");
                                    writer.WriteElementString("VerticalAlign", "Middle");
                                    writer.WriteElementString("HorizontalAlign", "Center");
                                    writer.WriteElementString("Margin", "0");
                                    writer.WriteElementString("Proportional", "False");
                                    writer.WriteElementString("Compress", "False");
                                    writer.WriteElementString("Freeze", "False");
                                    writer.WriteElementString("Orientation", "0");
                                    writer.WriteStartElement("TextContent");
                                    writer.WriteElementString("String", bmk_names[i].Replace("-", ""));
                                    writer.WriteStartElement("Font");
                                    writer.WriteAttributeString("RefersToID", "1");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("Color");
                                    writer.WriteAttributeString("RefersToID", "0");
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                }
                            }
                        }
                        //End Strip Blocks
                        writer.WriteEndElement();

                        //Text Attributes
                        writer.WriteStartElement("TextAttributes");
                        writer.WriteStartElement("Fonts");
                        writer.WriteStartElement("Font");
                        writer.WriteAttributeString("ID", "0");
                        writer.WriteElementString("FaceName", "smartFont");
                        writer.WriteElementString("Height", "2910");
                        writer.WriteElementString("Width", "1000");
                        writer.WriteElementString("Italic", "False");
                        writer.WriteElementString("Bold", "False");
                        writer.WriteElementString("Underline", "False");
                        writer.WriteElementString("StrikeOut", "False");
                        writer.WriteElementString("PitchAndFamily", "0x00000002");
                        writer.WriteElementString("CharSet", "1");
                        writer.WriteElementString("Plotter", "False");
                        writer.WriteEndElement();

                        writer.WriteStartElement("Font");
                        writer.WriteAttributeString("ID", "1");
                        writer.WriteElementString("FaceName", "smartFont");
                        writer.WriteElementString("Height", "3440");
                        writer.WriteElementString("Width", "1000");
                        writer.WriteElementString("Italic", "False");
                        writer.WriteElementString("Bold", "True");
                        writer.WriteElementString("Underline", "False");
                        writer.WriteElementString("StrikeOut", "False");
                        writer.WriteElementString("PitchAndFamily", "0x00000022");
                        writer.WriteElementString("OutPrecision", "0x00000003");
                        writer.WriteElementString("ClipPrecision", "0x00000002");
                        writer.WriteElementString("Quality", "0x00000001");
                        writer.WriteElementString("Plotter", "False");
                        writer.WriteEndElement();
                        writer.WriteStartElement("Font");
                        writer.WriteAttributeString("ID", "2");
                        writer.WriteElementString("FaceName", "smartFont");
                        writer.WriteElementString("Height", "2910");
                        writer.WriteElementString("Width", "1000");
                        writer.WriteElementString("Italic", "False");
                        writer.WriteElementString("Bold", "True");
                        writer.WriteElementString("Underline", "False");
                        writer.WriteElementString("StrikeOut", "False");
                        writer.WriteElementString("PitchAndFamily", "0x00000022");
                        writer.WriteElementString("OutPrecision", "0x00000003");
                        writer.WriteElementString("ClipPrecision", "0x00000002");
                        writer.WriteElementString("Quality", "0x00000001");
                        writer.WriteElementString("Plotter", "False");
                        writer.WriteEndElement();

                        writer.WriteEndElement();
                        writer.WriteStartElement("Colors");
                        writer.WriteStartElement("Color");
                        writer.WriteAttributeString("Format", "RGB");
                        writer.WriteAttributeString("ID", "0");
                        writer.WriteString("0x000000");
                        writer.WriteEndElement();

                        writer.WriteEndElement();
                        writer.WriteEndElement();


                        // device Section
                        writer.WriteStartElement("Device");

                        writer.WriteElementString("Type", "Printer");
                        writer.WriteElementString("DrawBorder", "False");
                        writer.WriteElementString("DrawBorderV", "False");
                        writer.WriteElementString("DrawContent", "True");
                        writer.WriteElementString("EmulateTrueType", "False");
                        writer.WriteElementString("DeviceID", "None");
                        writer.WriteStartElement("DevMode");
                        writer.WriteAttributeString("Encoding", "Hex");
                        writer.WriteString("5741474F20736D6172745072696E74657220284B6F7069652032290000000000010402079C0014050F25010001000001EA0A58026400010000012C01010001002C01010000005553455200720000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000200000003000000010000000001000000000000000000000000000000000000A40001004D44544E02000000050000002C01000000190000E40C000000000000B80B00007017000000000000DC05000058020000000000000000000000000000FEFF0000000000000000000000000000000000000000000000000000370C0000DA0C000010270000102700001027000010270000A00A0000C2060000B80600002C04000040010000D2000000180000000000102710271027000010270000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000006804010400000000CE180000020000000000000000000000020002000000120030000000000000000100000000000000AC0D0000B80B000000000000EFCDAB9001000000D00700000000000070C600000D00000000000000000000000000000000000000010000007373736764787004");
                        writer.WriteEndElement();
                        writer.WriteStartElement("DevNames");
                        writer.WriteAttributeString("Encoding", "Hex");
                        writer.WriteString("080012002F00000077696E73706F6F6C00005741474F20736D6172745072696E74657220284B6F7069652032290033555342303033000000000000000000000000000000000000");
                        writer.WriteEndElement();
                        writer.WriteEndElement();

                        /* foreach (WifiNetwork WifiNetwork in WifiNetworklist)
                         {
                             writer.WriteStartElement("WifiNetwork");

                             writer.WriteElementString("SSID", WifiNetwork.SSID);
                             writer.WriteElementString("Key", WifiNetwork.Key);
                             writer.WriteElementString("DHCPorSTATIC", WifiNetwork.DHCPorSTATIC);
                             writer.WriteElementString("StaticIP", WifiNetwork.StaticIP);

                             writer.WriteEndElement();
                         }*/

                        writer.WriteEndElement();
                        writer.WriteEndDocument();

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
        }


        /// <summary>
        /// Read the WifiNetwork file 
        /// </summary>
        /// <param name="WifiNetwork_list"></param>
        ///
        /*
              public static void Read_XML_WifiNetwork_File(List<WifiNetwork> WifiNetwork_list)
              {

                  try
                  {
                      // Console.WriteLine(xml_name);
                      if (File.Exists(xml_name))
                      {

                          XmlDocument xdoc = new XmlDocument();
                          xdoc.Load(xml_name);

                          foreach (XmlNode WifiNetwork in xdoc.SelectNodes("/WifiNetwork_List/*"))
                          {
                              if (WifiNetwork != null)
                              {
                                  WifiNetwork_list.Add(new WifiNetwork(WifiNetwork["SSID"].InnerText, WifiNetwork["Key"].InnerText, WifiNetwork["DHCPorSTATIC"].InnerText, WifiNetwork["StaticIP"].InnerText));
                                  // Console.WriteLine(WifiNetwork["SSID"].InnerText + WifiNetwork["Key"].InnerText + WifiNetwork["DHCPorSTATIC"].InnerText + WifiNetwork["StaticIP"].InnerText);
                              }

                          }


                      }
                  }
                  catch (Exception ex)
                  {
                      Console.WriteLine(ex.StackTrace);
                  }
              }

              public static void Change_XML_WifiNetwork_File(WifiNetwork WifiNetwork)
              {

                  try
                  {
                      // Console.WriteLine(xml_name);
                      if (File.Exists(xml_name))
                      {

                          XmlDocument xdoc = new XmlDocument();
                          xdoc.Load(xml_name);

                          XmlElement node = xdoc.CreateElement("WifiNetwork");
                          XmlElement SSID = xdoc.CreateElement("SSID");
                          SSID.InnerText = WifiNetwork.SSID;
                          XmlElement Key = xdoc.CreateElement("Key");
                          Key.InnerText = WifiNetwork.Key;
                          XmlElement DHCPorSTATIC = xdoc.CreateElement("DHCPorSTATIC");
                          DHCPorSTATIC.InnerText = WifiNetwork.DHCPorSTATIC;
                          XmlElement StaticIP = xdoc.CreateElement("StaticIP");
                          StaticIP.InnerText = WifiNetwork.StaticIP;
                          node.AppendChild(SSID);
                          node.AppendChild(Key);
                          node.AppendChild(DHCPorSTATIC);
                          node.AppendChild(StaticIP);
                          xdoc.DocumentElement.AppendChild(node);

                          xdoc.Save(xml_name);


                      }
                  }
                  catch (Exception ex)
                  {
                      Console.WriteLine(ex.StackTrace);
                  }
              }


              public class SelectNetwork
              {
                  string _netmsg;
                  bool _IsInList;
                  int _listindex;
                  string _SSID;

                  public SelectNetwork(string Netmsg, bool IsinList, int ListIndex, string ssid)
                  {
                      this._netmsg = Netmsg;
                      this._IsInList = IsinList;
                      this._listindex = ListIndex;
                      this._SSID = ssid;
                  }
                  public string netmsg { get { return _netmsg; } }
                  public bool IsInList { get { return _IsInList; } }
                  public int listindex { get { return _listindex; } }
                  public string SSID { get { return _SSID; } }
              }

              public class WifiNetwork
              {
                  string _ssid;
                  string _key;
                  string _DCHPorStatic;
                  string _staticIP;


                  public WifiNetwork(string ssid, string key, string DCHPorStatic, string staticIP)
                  {
                      this._ssid = ssid;
                      this._key = key;
                      this._DCHPorStatic = DCHPorStatic;
                      this._staticIP = staticIP;

                  }

                  public string SSID { get { return _ssid; } }
                  public string Key { get { return _key; } }
                  public string DHCPorSTATIC { get { return _DCHPorStatic; } }
                  public string StaticIP { get { return _staticIP; } }

              }*/

    }
}