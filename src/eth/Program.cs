// Program.cs
//
// Illustrates how to use the HP OTA to export the Test Lab hierarchy
// to a CSV-formatted file.
//
// Chris Trueman
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Mono.Options;

using TDAPIOLELib;

class Program
{
   private static string SIGNON_VERSION = "0.1.0";

   private string m_Url = "";
   private string m_Username = "";
   private string m_Password = "";
   private string m_Domain = "";
   private string m_Project = "";
   private string m_OutputPathname = "out.csv";
   private List<string> m_Labels = new List<string>();
   private bool m_ShowHelp = false;

   private TDConnectionClass m_Connection = null;
   private bool m_Silent = false;
   private StreamWriter m_Output = null;
   private int m_MaxLevel = -1;

   enum FactoryType
   {
      TSTest
   };

   private Dictionary<FactoryType, Dictionary<string, string>> m_FactoryLabelColumnMap = null;

   static void Main(string[] args)
   {
      using (new SimpleTimer())
      {
         Program p = new Program();

         p.Run(args);
      }
   }

   void Run(string[] args)
   {
      var options = new OptionSet()
      {
         {
            "s|server=", 
            "Quality Center server URL",
            (string v) => m_Url = v
         },
         {
            "u|username=",
            "Username",
            (string v) => m_Username = v
         },
         {
            "p|password=",
            "Password",
            (string v) => m_Password = v
         },
         {
            "d|domain=",
            "Domain",
            (string v) => m_Domain = v
         },
         {
            "j|project=",
            "Project",
            (string v) => m_Project = v
         },
         {
            "o|output=",
            "Output pathname",
            (string v) => m_OutputPathname = v
         },
         {
            "l|label=",
            "TSTest label to output",
            v => m_Labels.Add(v)
         },
         {
            "h|help", 
            "Program copyright, version and help",
            v => m_ShowHelp = (v != null)
         }
      };

      try
      {
         options.Parse(args);
      }
      catch (OptionException e)
      {
         Console.Write("eth: ");
         Console.WriteLine(e.Message);
         Console.WriteLine("Try 'eth.exe --help' for more information.");

         return;
      }

      if (m_ShowHelp || !ValidateArgs())
      {
         ShowHelp(options);

         return;
      }

      try
      {
         InitializeOutput();
         InitializeConnection();
         LoadFactoryTypesLabelColumns();
         ExportTestSetTree();
         FinalizeConnection();
         FinalizeOutput();
      }
      catch (Exception e)
      {
         Console.WriteLine();
         Console.Write("eth: ");
         Console.WriteLine(e.Message);
         Console.WriteLine(e.StackTrace);
      }
   }

   bool ValidateArgs()
   {
      if (string.IsNullOrEmpty(m_Url))
         return false;

      if (string.IsNullOrEmpty(m_Username))
         return false;

      if (string.IsNullOrEmpty(m_Domain))
         return false;

      if (string.IsNullOrEmpty(m_Project))
         return false;

      return true;
   }

   void ShowHelp(OptionSet options)
   {
      Console.WriteLine("ETH.EXE - EXPORT TEST HIERARCHY - v{0}", SIGNON_VERSION);
      Console.WriteLine("COPYRIGHT 2014 CHRIS TRUEMAN. ALL RIGHTS RESERVED.");
      Console.WriteLine();
      Console.WriteLine("Options:");
      options.WriteOptionDescriptions(Console.Out);
   }

   void InitializeOutput()
   {
      m_Output = new StreamWriter(m_OutputPathname);
   }

   void FinalizeOutput()
   {
      m_Output.Close();
   }

   void InitializeConnection()
   {
      bool silent = SetSilent(false);

      Log("Connect server: ");

      m_Connection = new TDConnectionClass();

      if (null == m_Connection)
         throw new Exception("can't create TDConnection object.");

      m_Connection._DIProgressEvents_Event_OnProgress += new _DIProgressEvents_OnProgressEventHandler(OnProgress);
      m_Connection.InitConnectionEx(m_Url);

      Log(" done.\n");

      Log("Login: ");
      m_Connection.Login(m_Username, m_Password);
      Log(" done.\n");

      Log("Connect project: ");
      m_Connection.Connect(m_Domain, m_Project);
      Log(" done.\n");

      Log("Ignore HTML formatting: ");
      m_Connection.IgnoreHtmlFormat = true;
      Log(" done.\n");

      SetSilent(silent);
   }

   void FinalizeConnection()
   {
      bool silent = SetSilent(false);

      Log("Disconnect: ");
      m_Connection.DisconnectProject();
      Log(" done.\n");

      SetSilent(silent);
   }

   void ExportTestSetTree()
   {
      bool silent = SetSilent(true);

      TestSetTreeManager test_set_tree_manager = (TestSetTreeManager)m_Connection.TestSetTreeManager;
      TestSetFolder root = (TestSetFolder)test_set_tree_manager.Root;
      Stack<string> elements = new Stack<string>();

      //
      // Walk the hierarchy to discover its max level.
      //

      ExportTestSetTree_TestSetFolder(root, 0, elements, false);

      Log("Max level: {0}\n", m_MaxLevel);

      //
      // Output hierarchy.
      //

      ExportTestSetTree_TestSetFolder(root, 0, elements, true);

      SetSilent(silent);
   }

   void ExportTestSetTree_TestSetFolder(TestSetFolder folder, int level, Stack<string> elements, bool writeOutput)
   {
      List kids = folder.NewList();

      foreach (TestSetFolder kid in kids)
      {
         string name = kid.Name;

         Log("{0}{1}\n", Indent(level), name);

         elements.Push(name);

         ExportTestSetTree_TestSets(kid, level + 1, elements, writeOutput);
         ExportTestSetTree_TestSetFolder(kid, level + 1, elements, writeOutput);

         elements.Pop();
      }
   }

   void ExportTestSetTree_TestSets(TestSetFolder folder, int level, Stack<string> elements, bool writeOutput)
   {
      TestSetFactory test_set_factory = (TestSetFactory)folder.TestSetFactory;
      List test_sets = test_set_factory.NewList("");

      foreach (TestSet test_set in test_sets)
      {
         string name = test_set.Name;

         Log("{0}{1}\n", Indent(level), name);

         elements.Push(name);

         ExportTestSetTree_TSTests(test_set, level + 1, elements, writeOutput);

         elements.Pop();
      }
   }

   void ExportTestSetTree_TSTests(TestSet testSet, int level, Stack<string> elements, bool writeOutput)
   {
      if (level > m_MaxLevel)
         m_MaxLevel = level;

      TSTestFactory ts_test_factory = (TSTestFactory)testSet.TSTestFactory;
      List ts_tests = ts_test_factory.NewList("");

      foreach (TSTest ts_test in ts_tests)
      {
         string name = ts_test.Name;

         Log("{0}{1}:\n", Indent(level), name);

         elements.Push(name);

         IBaseField thing = (IBaseField)ts_test;
         List<string> values = new List<string>();
         foreach (string label in m_Labels)
         {
            string value = "";

            Log("{0}{1} => ", Indent(level + 1), label);

            if (m_FactoryLabelColumnMap.ContainsKey(FactoryType.TSTest))
            {
               Dictionary<string, string> label_column_map = m_FactoryLabelColumnMap[FactoryType.TSTest];

               if (label_column_map.ContainsKey(label))
               {
                  string column = label_column_map[label];

                  try
                  {
                     object field = thing[column];

                     if (null != field)
                        value = field.ToString();
                  }
                  catch (System.Runtime.InteropServices.COMException /*e*/)
                  {
                     // Field doesn't exist. Ignore.

                     /*
                      *  (0x8004051C): Invalid field name < field_name >
                      */
                  }
               }
            }

            values.Add(value);

            Log("{0}\n", value);
         }

         if (writeOutput)
            WriteOutput(elements, values);

         elements.Pop();
      }
   }

   void WriteOutput(Stack<string> elements, List<string> values)
   {
      Stack<string> ordered_elements = new Stack<string>(elements);
      bool first = true;

      // Tree items.

      foreach (string ordered_element in ordered_elements)
      {
         if (!first)
            m_Output.Write(",");

         m_Output.Write("\"{0}\"", ordered_element);
         first = false;
      }

      // Pad output to m_MaxLevel.

      for (int iii = ordered_elements.Count; iii <= m_MaxLevel; ++iii)
      {
         m_Output.Write(",");
      }

      // And finally the leaf values.

      foreach (string value in values)
      {
         m_Output.Write(",\"{0}\"", value);
      }

      m_Output.WriteLine();
   }

   void LoadFactoryTypesLabelColumns()
   {
      bool silent = SetSilent(true);

      Log("Load factory label columns map:\n");
      LoadFactoryTypeLabelColumns(FactoryType.TSTest, (IBaseFactory)m_Connection.TSTestFactory);
      Log("...done.\n");

      SetSilent(silent);
   }

   void LoadFactoryTypeLabelColumns(FactoryType factoryType, IBaseFactory factory)
   {
      if (null == m_FactoryLabelColumnMap)
         m_FactoryLabelColumnMap = new Dictionary<FactoryType, Dictionary<string, string>>();

      if (m_FactoryLabelColumnMap.ContainsKey(factoryType))
         return;

      Log("   {0}:\n", factoryType.ToString());

      Dictionary<string, string> label_column_map = new Dictionary<string, string>();
      List fields = factory.Fields;

      foreach (TDField field in fields)
      {
         FieldProperty field_property = (FieldProperty)field.Property;

         if (field_property.IsActive)
         {
            string label = field_property.UserLabel;
            string column = field_property.DBColumnName;

            if (!string.IsNullOrEmpty(label) &&
               !label_column_map.ContainsKey(label))
            {
               label_column_map.Add(label, column);
               Log("      {0} => {1}\n", label, column);
            }
         }
      }

      m_FactoryLabelColumnMap.Add(factoryType, label_column_map);
   }

   string Indent(int count)
   {
      return "".PadLeft(count * 2);
   }

   bool SetSilent(bool silent)
   {
      bool old_silent = m_Silent;

      m_Silent = silent;

      return old_silent;
   }

   void OnProgress(int current, int total, string message)
   {
      if (!m_Silent)
         Log(".");
   }

   void Log(string format, params object[] args)
   {
      Console.Write(format, args);
   }
}
