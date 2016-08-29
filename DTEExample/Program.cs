using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using EnvDTE;
using EnvDTE80;
using EnvDTE90;
using EnvDTE100;

using System.IO;

namespace DTEExample
{
    class Program
    {
        static void Main(string[] args)
        {
            //
            // This project uses EnvDTE to generate code in Visual Studio and save the results, it contains a ton of examples
            //

            //----------------------------------------------------------------------------------------------------------
            //
            // Layouts
            //
            // DTE -> Solution -> Projects -> Project -> ProjectItems -> ProjectItem -> FileCodeModel -> CodeElements
            //
            //
            //
            // DTE strucutre
            //
            // DTE
            //  Solution
            //   Projects
            //    Project
            //     ProjectItems
            //      ProjectItem
            //       CodeNamespace
            //       
            //----------------------------------------------------------------------------------------------------------

            // Initialize (almost) everything

            // Name stuff
            string SolutionName = "SolutionName";
            string ProjectName = "ProjectName";

            string ThisdllDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // replace the below line with a modification of the above
            string parentDir = Directory.GetParent(ThisdllDirectory).Parent.FullName;

            string CurrentDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            // template directories
            string TemplatesDirectory = CurrentDirectory + @"\Templates\";
            string ProjectTemplate = TemplatesDirectory + @"ProjectEmpty\ProjectEmpty.vstemplate";  //C:\GoogleDrive\TAMagicDevelopment\DTEExample\DTEExample\Templates\ProjectEmpty\ProjectEmpty.vstemplate
            string ItemTemplate = TemplatesDirectory + @"ItemClass\ItemClass.vstemplate";           //C:\GoogleDrive\TAMagicDevelopment\DTEExample\DTEExample\Templates\ItemClass\ItemClass.vstemplate
            string ItemTemplateEmpty = TemplatesDirectory + @"ItemClassEmpty\ItemClassEmpty.vstemplate";  //C:\GoogleDrive\TAMagicDevelopment\DTEExample\DTEExample\Templates\ItemClassEmpty\ItemClassEmpty.vstemplate

            // save directories
            string SolutionRootDirectory = @"C:\Temp\";
            string SolutionDirectory = SolutionRootDirectory + SolutionName + @"\";
            string ProjectDirectory = SolutionDirectory + ProjectName + @"\";

            // For ItemName2 and ItemName3 examples below
            List<string> DirectoryList = new List<string>();
            string FirstLevelSubDirectory = @"FirstLvlSubDir\";                             // added to the root of the ProjectDirectory later in code
            string SecondLevelSubDirectory = FirstLevelSubDirectory + @"SecondLvlSubDir\";  // added to the root of the ProjectDirectory later in code

            //-------------------------

            EnvDTE.Project currentProject = null;
            EnvDTE.ProjectItem currentItem = null;

            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------

            // display the directories for grins and giggles
            
            Console.WriteLine("TemplatesDirectory > " + TemplatesDirectory);
            Console.WriteLine("ProjectTemplate > " + ProjectTemplate);
            Console.WriteLine("ItemTemplate > " + ItemTemplate);
            Console.WriteLine("SolutionRootDirectory > " + SolutionRootDirectory);
            Console.WriteLine("SolutionDirectory > " + SolutionDirectory);
            Console.WriteLine("ProjectDirectory > " + ProjectDirectory);
            // subdirectories added
            Console.WriteLine("FirstLevelSubDirectory > " + FirstLevelSubDirectory);
            Console.WriteLine("SecondLevelSubDirectory > " + SecondLevelSubDirectory);
            Console.WriteLine("\n\nThisdllDirectory parentDir> {0}", parentDir);
            Console.WriteLine("CurrentDirectory > " + CurrentDirectory);
            //--------------------------------------------------------------------
            //--------------------------------------------------------------------
            //--------------------------------------------------------------------


            FileSystemMethods.DeleteFilesAndDirectories(SolutionDirectory); // remove the old version so this does not exception out

            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------

            //
            // Create Everything
            //

            // DTE -> Solution -> Projects -> Project -> ProjectItems -> ProjectItem -> FileCodeModel -> CodeElements

            //DTE - create DTEBase.dte - done this way so multiple solutions can be added to a list (future project) to ease the verboseness of coding DTE
            EnvDTEBase DTEBase = new EnvDTEBase();  // instantiate a new EnvDTE
            EnvDTEBase.InitializeDTEBase(DTEBase);             // Now initialize it

            //-------------------------------------------------------------------------------------------------

            // Solution - create the solution
            DTEBase.dte.Solution.Create(SolutionDirectory, SolutionName);

            //-------------------------------------------------------------------------------------------------

            //  Projects
            //   Project

            //-------------------------------------------------------------------------------------------------
            //
            //   Project - create a project from a template
            //
            DTEBase.dte.Solution.AddFromTemplate(ProjectTemplate, SolutionDirectory + ProjectName + @"\", ProjectName); // this returns null hence the next line
            currentProject = DTEBase.dte.Solution.Projects.GetProjectByName(ProjectName); // (GetProjectByName), using the extension class (EnvDTEExtensions) see below...


            //-------------------------------------------------------------------------------------------------

            //    ProjectItems
            //     ProjectItem

            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------
            //-------------------------------------------------------------------------------------------------
            //
            //
            //     ProjectItem - (ItemName1) create a class in the project root from a template
            //
            //
            currentItem = currentProject.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName1"); // returns null to "currentItem" hence the next line
            currentItem = currentProject.GetItemByName("ItemName1" + ".cs"); // (GetItemByName), using the extension class (EnvDTEExtensions) see below...


            //
            // (ItemName1) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);

            //-------------------------------------------------------------------------------------------------
            //
            //
            //     ProjectItem - (ItemName2) create a class in a sub directory of the project root
            //
            //

            DirectoryList.Clear();
            DirectoryList.ConvertDirectoryStringToList(FirstLevelSubDirectory);

            Console.WriteLine("++------Showing directories from root");
            foreach (string str in DirectoryList)
            {
                Console.WriteLine("{0}", str);
            }
            Console.WriteLine("++------");

            //-------

            currentItem = currentProject.AddFolders(DirectoryList);

            currentItem.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName2");
            currentItem = EnvDTEExtensions.FindProjectItemInProject(currentProject, "ItemName2" + ".cs", true);


            //
            // (ItemName2) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            //     ProjectItem - (ItemName2a) create a class, With a namespace not reliant on the directory structure
            //
            //

            currentProject.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName2a");
            currentItem = EnvDTEExtensions.FindProjectItemInProject(currentProject, "ItemName2a" + ".cs", true);


            CodeNamespace cnsIN2a = currentItem.FileCodeModel.AddNamespace("ItemName2aNameSpaceName", -1);  // 0 = add a name space at the beginning, -1 at end, 0 will add it after the //{{End Using}} comment
            CodeClass ccIN2a = cnsIN2a.AddClass("ItemName2aClassName", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);



            //
            // (ItemName2) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            //     ProjectItem - (ItemName3) create a class in a sub, sub directory of the project root, 
            //
            //

            DirectoryList.Clear();
            DirectoryList.ConvertDirectoryStringToList(SecondLevelSubDirectory);

            Console.WriteLine("++------Showing directories from root");
            foreach (string str in DirectoryList)
            {
                Console.WriteLine("{0}", str);
            }
            Console.WriteLine("++------");

            //------- 

            currentItem = currentProject.AddFolders(DirectoryList);

            currentItem.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName3");
            currentItem = EnvDTEExtensions.FindProjectItemInProject(currentProject, "ItemName3" + ".cs", true); // works on any file in a sub dir

            //
            // (ItemName3) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);

            //-------------------------------------------------------------------------------------------------
            //
            //
            //     ProjectItem - (ItemName4) Edit and add Namespaces, Methods, Fields, etc in root directory
            //
            //
            currentProject.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName4");
            currentItem = currentProject.GetItemByName("ItemName4" + ".cs");

            //
            // Inserting things into a file/Item
            //

            CodeNamespace cnsIN4 = currentItem.FileCodeModel.AddNamespace("ItemName4NameSpaceName", -1);  // 0 = add a name space at the beginning, -1 at end, 0 will add it after the //{{End Using}} comment
            CodeClass ccIN4 = cnsIN4.AddClass("ItemName4ClassName", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);
            CodeVariable cvIN4 = ccIN4.AddVariable("IntVariableName", vsCMTypeRef.vsCMTypeRefInt, Type.Missing, vsCMAccess.vsCMAccessPublic, Type.Missing);
            // functionName, functionType, ReturnType, 0 beginning -1 end else n position, access level, ??? - go look this up)
            CodeFunction cfIN4 = ccIN4.AddFunction("FunctionName", vsCMFunction.vsCMFunctionFunction, vsCMTypeRef.vsCMTypeRefInt, -1, vsCMAccess.vsCMAccessPublic, Type.Missing);
            //name, type, position (-1 end 0 begin n = location)
            cfIN4.AddParameter("Parameter1", vsCMTypeRef.vsCMTypeRefByte, -1); // 1st
            cfIN4.AddParameter("Parameter2", vsCMTypeRef.vsCMTypeRefByte, -1); // 2cnd
            cfIN4.AddParameter("ParameterMiddle", vsCMTypeRef.vsCMTypeRefByte, 1); // now 2cnd and previous one 3rd

            // now lets add some code to the function
            // https://social.msdn.microsoft.com/Forums/vstudio/en-US/f23d6bfa-2cd8-488f-aa3a-07970e40b247/envdte-projectitem-filecodemodel-how-to-create-method-body?forum=vsx
            // http://stackoverflow.com/questions/13376944/programmatically-add-function-to-existing-c-file-with-envdte
            TextPoint tpIN4 = cfIN4.GetStartPoint(vsCMPart.vsCMPartBody);
            EditPoint epIN4 = tpIN4.CreateEditPoint();
            //ep.LineDown(3); //goes down 3 lines, does not add lines

            epIN4.Indent(null, 0); // this is relative to the current indentation
            epIN4.Insert("//Holy bat crap man!!! dyslexic batman-robin joke ;-)");
            epIN4.Insert("\n"); // new line

            epIN4.Indent(null, 2);
            epIN4.Insert("int myVar = 1;");
            epIN4.Insert("\n"); // new line

            epIN4.Indent(null, 3);
            epIN4.Insert("myVar = myVar * 100;");
            epIN4.Insert("\n"); // new line

            epIN4.Indent(null, 3);
            Console.WriteLine("currentItem.Collection.Count = {0}", currentItem.Collection.Count);

            //
            // (ItemName4) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            // ProjectItem - (ItemName5) Edit and add Namespaces, Methods, Fields, etc in root directory
            // now lets create some "code" ie fill the thing with comments, where we want it for the trade engine build
            //
            //

            //CodeNamespace cnEdit = currentItem
            currentProject.ProjectItems.AddFromTemplate(ItemTemplateEmpty, "ItemName5");
            currentItem = currentProject.GetItemByName("ItemName5" + ".cs");

            CodeNamespace cnsIN5 = currentItem.FileCodeModel.AddNamespace("NameSpaceName", -1); // 0 = add a name space at the beginning, -1 at end
            CodeClass ccIN5 = cnsIN5.AddClass("NewClassName", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);


            TextPoint tpIN5 = ccIN5.GetEndPoint(vsCMPart.vsCMPartBody);// GetStartPoint(vsCMPart.vsCMPartBody);
            EditPoint epIN5 = tpIN5.CreateEditPoint();

            // formating the entire document

            epIN5.Indent(null, 0); // this is relative to the current indentation
            epIN5.Insert("\n"); // new line

            epIN5.Indent(null, 1); // this is not
            epIN5.Insert("// 1 This should not go at the end of the file!!!");
            epIN5.Insert("\n"); // new line

            epIN5.Indent(null, 2); // this is not
            epIN5.Insert("// 2 This should go next");
            epIN5.Insert("\n"); // new line

            epIN5.Indent(null, 4); // this is not
            epIN5.Insert("                                      // 2b experimenting with spaces");
            epIN5.Insert("\n"); // new line

            epIN5.Indent(null, 5);
            epIN5.Insert("// 3 this should go last");

            for (int i = 0; i < 50; i++)
            {
                for (int j = 0; j < i; j++)
                {
                    epIN5.Insert("\t");
                }
                epIN5.Insert("//HelloWorld" + i.ToString() + "\n");
            }

            //
            // (ItemName5) lets demo the reformat routine by itself
            //
            currentItem.Reformat(); // demo to straignten up after the for loop above

            //
            // (ItemName5) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it ... yeah I know the reformat is already done ... 
            //
            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            // ProjectItem - (ItemName6) locate the line number some text is at and insert above it
            //
            //
            Console.WriteLine("ProjectItem - (ItemName6) locate the line number some text is at and insert above it");

            currentProject.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName6"); // use the template with the comment sections for inserting code into
            currentItem = currentProject.GetItemByName("ItemName6" + ".cs");

            // create some stuff to see if we can find it
            CodeNamespace cns6Edit = currentItem.FileCodeModel.AddNamespace("NameSpaceName", -1); // 0 = add a name space at the beginning, -1 at end
            CodeClass cc6Edit = cns6Edit.AddClass("NewClassName6a", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);
            cc6Edit = cns6Edit.AddClass("NewClassName6b", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);
            cc6Edit = cns6Edit.AddClass("NewClassName6c", Type.Missing, Type.Missing, Type.Missing, vsCMAccess.vsCMAccessPublic);


            Console.WriteLine("scan throught the code and describe the elements");
            EnvDTE.CodeElements codeElem6 = currentItem.FileCodeModel.CodeElements;
            foreach (EnvDTE.CodeElement ce in codeElem6)
            {
                Console.WriteLine(ce.Kind + " <> " + "" + " <> ");
                foreach (CodeElement childElement in ce.Children)
                {
                    EnvDTEExtensions.ExamineCodeElement(childElement, 1);
                }
            }

            //---------------------------- 
            // Here is where things get interesting!

            Console.WriteLine("Lets try a different way...");

            TextSelection objSel = currentItem.Document.Selection;
            objSel.GotoLine(1, false); // (lineNum, select bool)
            objSel.Insert("//I am at the top of the document!!\n\n//Hello World!!!!!\n\n", 0);

            objSel.StartOfDocument(true);
            objSel.Insert("//objSel.StartOfDocument(true) test\n\n", 0);

            Console.WriteLine("Now we have access to the entire document!!!!");

            //----------------

            // go to the first line
            objSel.GotoLine(1, true);
            // write the line out
            Console.WriteLine(objSel.Text);
            Console.WriteLine(objSel.LastLine());

            Console.WriteLine("Now we can scan for text and mark locations!!!!");

            //-----------------

            Console.WriteLine(" objSel.GetLineNumberContaining(''//{{End Properties}}'') = {0}", objSel.GetLineNumberContaining("//{{End Properties}}"));

            //objSel.DisplayEntireDocument();

            // single line insert
            objSel.InsertAboveThisText(ClassInsertTokens.EndMethods, "//This is the time for all good men to come to the aid of their country");

            objSel.InsertBelowThisText(ClassInsertTokens.BeginClasses, "//This is the time for all good men to come to the aid of their country");

            //--------

            // multiline insert
            string multiLineInsert =
                "public void bob()\n" +
                "{\n" +
                "int i = 0;\n" +
                "i++;\n" +
                "Console.WriteLine(\"{0}\",i);\n" +
                "}";
            objSel.InsertAboveThisText(ClassInsertTokens.EndMethods, multiLineInsert);

            // insert a list of strings
            List<string> los = new List<string>() { "public void LosInsertTest()", "{", "//This is a comment", "int x = 42; // meaning of life the universe and everything", "Console.WriteLine(\"The question is ... What is 7 times eight \");", "Console.WriteLine(x);", "}" };

            // insert at the end
            objSel.InsertAboveThisText(ClassInsertTokens.EndMethods, los);

            // insert at the beginning
            


            //objSel.InsertBelowThisText(ClassInsertTokens.BeginClasses, los);
            //objSel.DisplayEntireDocument();

            currentItem.InsertBelowThisText(ClassInsertTokens.BeginClasses, los);
            currentItem.Reformat();
            currentItem.DisplayEntireDocument();
            
            //
            // (ItemName6) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //


            currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            // ProjectItem - (ItemName7) Repeat the above (ItemName6) with CurrentItem
            //
            // using, Fields, Properties, Constructors, Methods, Enums, Destructors, indexers, events, classes, 
            // static classes, constants, operators, delegates, intefaces structs
            //
            //

            Console.WriteLine("ProjectItem - (ItemName7) locate the line number some text is at and insert above it");

            currentProject.ProjectItems.AddFromTemplate(ItemTemplate, "ItemName7"); // use the template with the comment sections for inserting code into
            //currentItem = currentProject.GetItemByName("ItemName7" + ".cs");
            //TextSelection textSel = currentItem.Document.Selection;
            currentItem = EnvDTEExtensions.FindProjectItemInProject(currentProject, "ItemName7" + ".cs", true); // works on any file in a sub dir

            string multiLine =
                "public void bob()\n" +
                "{\n" +
                "int i = 0;\n" +
                "i++;\n" +
                "Console.WriteLine(\"{0}\",i);\n" +
                "}";
            List<string> los7 = new List<string>() 
            { 
                "public void LosInsertTest()", 
                "{", "//This is a comment", 
                "int x = 42; // meaning of life the universe and everything", 
                "Console.WriteLine(\"The question is ... What is 7 times eight \");", 
                "Console.WriteLine(x);", 
                "}" 
            };

            currentItem.InsertBelowThisText(ClassInsertTokens.BeginConstructors, multiLine);

            currentItem.InsertAboveThisText(ClassInsertTokens.EndFields, "//currentItem.InsertAboveThisText(ClassInsertTokens.EndFields...");
            currentItem.InsertBelowThisText(ClassInsertTokens.BeginFields, "//currentItem.InsertBelowThisText(ClassInsertTokens.BeginFields...");

            currentItem.InsertAboveThisText(ClassInsertTokens.EndMethods, los7);
            currentItem.InsertBelowThisText(ClassInsertTokens.BeginFields, los7);


            currentItem.Reformat();
            
            //currentItem.DisplayEntireDocument();

            //
            // (ItemName7) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //

            //currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            // ProjectItem - (ItemName8) Now lets extend a class after the fact
            //
            //

            //
            // (ItemName8) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //

            //currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);


            //-------------------------------------------------------------------------------------------------
            //
            //
            // ProjectItem - (ItemName9) add and implement an interface on a class after the fact
            //
            //

            //
            // (ItemName9) now lets reformat (ctrl-K-D), Save, and close the window that is being edited in and save it
            //

            //currentItem.SaveAndClose();  //replaces - currentItem.Document.Close(vsSaveChanges.vsSaveChangesYes);



            //-------------------------------------------------------------------------------------------------
            //
            //
            //
            //
            //



            //-------------------------------------------------------------------------------------------------

            //
            // Loop through everything and display it. 
            //

            //DTE
            // Solution
            //  Projects
            //   Project
            //    ProjectItems
            //     ProjectItem

            Console.WriteLine("Press Enter To Close Window");
            Console.ReadLine();

            EnvDTEBase.SaveAndCloseDTE(DTEBase); // must be run to close the visual studio window and save the solution, and also close devenv.exe so you do not end up having 20 copies running in the back ground (see task manager).
        }
    }


    public class InsertTokens
    {
        // TODO each one of these "Begin/Endxxxxx" needs a "public class xxxxxCode" class associates with it like "public class ClassCode {}"

        public const string BeginClasses = "//{{Begin Classes}}";
        public const string EndClasses = "//{{End Classes}}";
        public const string BeginConstants = "//{{Begin Constants}}";
        public const string EndConstants = "//{{End Constants}}";
        public const string BeginConstructors = "//{{Begin Constructors}}";
        public const string EndConstructors = "//{{End Constructors}}";
        public const string BeginDelegates = "//{{Begin Delegates}}";
        public const string EndDelegates = "//{{End Delegates}}";
        public const string BeginDestructor = "//{{Begin Destructor}}";
        public const string EndDestructor = "//{{End Destructor}}";
        public const string BeginEnums = "//{{Begin Enums}}";
        public const string EndEnums = "//{{End Enums}}";
        public const string BeginEvents = "//{{Begin Events}}";
        public const string EndEvents = "//{{End Events}}";
        public const string BeginFields = "//{{Begin Fields}}";
        public const string EndFields = "//{{End Fields}}";
        public const string BeginIndexers = "//{{Begin Indexers}}";
        public const string EndIndexers = "//{{End Indexers}}";
        public const string BeginInterfaces = "//{{Begin Interfaces}}";
        public const string EndInterfaces = "//{{End Interfaces}}";
        public const string BeginMethods = "//{{Begin Methods}}";
        public const string EndMethods = "//{{End Methods}}";
        public const string BeginProperties = "//{{Begin Properties}}";
        public const string EndProperies = "//{{End Properties}}";
        public const string BeginOperators = "//{{Begin Operators}}";
        public const string EndOperators = "//{{End Operators}}";
        public const string BeginStaticClasses = "//{{Begin StaticClasses}}";
        public const string EndStaticClasses = "//{{End StaticClasses}}";
        public const string BeginStructs = "//{{Begin Structs}}";
        public const string EndStructs = "//{{End Structs}}";
        public const string BeginUsing = "//{{Begin Using}}";
        public const string EndUsing = "//{{End Using}}";
    }
    public class ReplaceTokens
    {
        // TODO move all the replace tokens up here and do what was done with InsertTokens >>--> ClassInsertTokens
    }
    public class ClassInsertTokens
    {
        public const string BeginClasses = InsertTokens.BeginClasses;
        public const string EndClasses = InsertTokens.EndClasses;
        public const string BeginConstants = InsertTokens.BeginConstants;
        public const string EndConstants = InsertTokens.EndConstants;
        public const string BeginConstructors = InsertTokens.BeginConstructors;
        public const string EndConstructors = InsertTokens.EndConstructors;
        public const string BeginDelegates = InsertTokens.BeginDelegates;
        public const string EndDelegates = InsertTokens.EndDelegates;
        public const string BeginDestructor = InsertTokens.BeginDestructor;
        public const string EndDestructor = InsertTokens.EndDestructor;
        public const string BeginEnums = InsertTokens.BeginEnums;
        public const string EndEnums = InsertTokens.EndEnums;
        public const string BeginEvents = InsertTokens.BeginEvents;
        public const string EndEvents = InsertTokens.EndEvents;
        public const string BeginFields = InsertTokens.BeginFields;
        public const string EndFields = InsertTokens.EndFields;
        public const string BeginIndexers = InsertTokens.BeginIndexers;
        public const string EndIndexers = InsertTokens.EndIndexers;
        public const string BeginInterfaces = InsertTokens.BeginInterfaces;
        public const string EndInterfaces = InsertTokens.EndInterfaces;
        public const string BeginMethods = InsertTokens.BeginMethods;
        public const string EndMethods = InsertTokens.EndMethods;
        public const string BeginProperties = InsertTokens.BeginProperties;
        public const string EndProperies = InsertTokens.EndProperies;
        public const string BeginOperators = InsertTokens.BeginOperators;
        public const string EndOperators = InsertTokens.EndOperators;
        public const string BeginStaticClasses = InsertTokens.BeginStaticClasses;
        public const string EndStaticClasses = InsertTokens.EndStaticClasses;
        public const string BeginStructs = InsertTokens.BeginStructs;
        public const string EndStructs = InsertTokens.EndStructs;
        public const string BeginUsing = InsertTokens.BeginUsing;
        public const string EndUsing = InsertTokens.EndUsing;
    }
    public class ClassReplaceTokens
    {
        // TODO Move to Namespace Tokens
        public const string NamespaceName = "{{NamespaceName}}";

        //
        public const string ClassAccessModifier = "{{ClassAccessModifier}}";
        public const string ClassOtherModifier = "{{ClassOtherModifier}}";
        public const string ClassName = "{{ClassName}}";
        public const string ClassExtends = "{{ClassExtends}}";
        public const string ClassExtendsInterface = "{{ClassExtendsInterface}}";

        // TODO Editing Tokens
        public const string crlf = "{{crlf}}";
        public const string tab = "{{tab}}"; // not needed because of "public static void Reformat(this EnvDTE.ProjectItem projItem)" 

        // TODO figure out the classes layout for the following
        // TODO MethodReplaceTokens ???
        // TODO OtherReplaceTokens ????
        // TODO OtherTokensFor Fields, destructors, etc ????
    }

    public class ClassCode // contains all the settings needed to create a class, methods for the header, and return
    {


    }

    public static class ClassEditMethods // rename this ClassEdit????
    {
        // TODO ClassClean(this ???? ClassCS) {} // deletes all the unused ClassReplaceTokens and removes the ClassInsertTokens

        // TODO ReplaceToken(this EnvDTE.ProjectItem currentItem, ClassReplaceTokens searchText, string replacementText) {}

        // TODO Create AddNewClass

        public static int GetLineNumberContaining(this EnvDTE.TextSelection textSel, ClassInsertTokens searchString)
        {
            return TextSelectionExtensions.GetLineNumberContaining(textSel, searchString.ToString());
        }
        public static int GetLineNumberContaining(this EnvDTE.ProjectItem currentItem, ClassInsertTokens searchString)
        {
            return TextSelectionExtensions.GetLineNumberContaining(currentItem, searchString.ToString());
        }
        public static void InsertAboveThisText(this EnvDTE.TextSelection textSel, ClassInsertTokens searchString, string insertString)
        {
            TextSelectionExtensions.InsertAboveThisText(textSel, searchString.ToString(), insertString);
        }
        public static void InsertAboveThisText(this EnvDTE.ProjectItem currentItem, ClassInsertTokens searchString, string insertString)
        {
            TextSelectionExtensions.InsertAboveThisText(currentItem, searchString.ToString(), insertString);
        }
        public static void InsertAboveThisText(this EnvDTE.TextSelection textSel, ClassInsertTokens searchString, List<string> listOfStrings)
        {
            TextSelectionExtensions.InsertAboveThisText(textSel, searchString.ToString(), listOfStrings);
        }
        public static void InsertAboveThisText(this EnvDTE.ProjectItem currentItem, ClassInsertTokens searchString, List<string> listOfStrings)
        {
            TextSelectionExtensions.InsertAboveThisText(currentItem, searchString.ToString(), listOfStrings);
        }
        public static void InsertBelowThisText(this EnvDTE.TextSelection textSel, ClassInsertTokens searchString, string insertString)
        {
            TextSelectionExtensions.InsertBelowThisText(textSel, searchString.ToString(), insertString);
        }
        public static void InsertBelowThisText(this EnvDTE.ProjectItem currentItem, ClassInsertTokens searchString, string insertString)
        {
            TextSelectionExtensions.InsertBelowThisText(currentItem, searchString.ToString(), insertString);
        }
        public static void InsertBelowThisText(this EnvDTE.TextSelection textSel, ClassInsertTokens searchString, List<string> listOfStrings)
        {
            TextSelectionExtensions.InsertBelowThisText(textSel, searchString.ToString(), listOfStrings);
        }
        public static void InsertBelowThisText(this EnvDTE.ProjectItem currentItem, ClassInsertTokens searchString, List<string> listOfStrings)
        {
            TextSelectionExtensions.InsertBelowThisText(currentItem, searchString.ToString(), listOfStrings);
        }

        //(this EnvDTE.TextSelection textSel, List<string> listOfStrings)
        //(this EnvDTE.ProjectItem currentItem, List<string> listOfStrings)

        public static void InsertClass() { }
        public static void InsertConstant() { }
        public static void InsertConstructor() { }
        public static void InsertDelegate() { }
        public static void InsertDestructor() { }
        public static void InsertEnum() { }
        public static void InsertEvent() { }
        public static void InsertField() { }
        public static void InsertIndexer() { }
        public static void InsertInterface() { }
        public static void InsertMethod() { }
        public static void InsertOperator() { }
        public static void InsertProperty() { }
        public static void InsertStaticClass() { }
        public static void InsertStruct() { }		
        public static void InsertUsingStatement() { }
        public static void InsertReference() // TODO might want to move this one to the project level, or have it at both levels, with this calling the project level  
        {
            ////browseUrl is either the File Path or the Strong Name
            ////(System.Configuration, Version=2.0.0.0, Culture=neutral,
            ////                       PublicKeyToken=B03F5F7F11D50A3A)
            //// http://www.codeproject.com/Articles/36219/Exploring-EnvDTE
            ////
            //
            //using EnvDTE;//Need!
            //using EnvDTE80;
            //using EnvDTE90;
            //
            //using VsLangProj;//Need!
            //using VsLangProj2;
            //using VsLangProj80;
            //
            //using VSWebSite.Interop;//Need!
            //using VSWebSite.Interop90;
            //
            //public void AddReference(string referenceStrIdentity, string browseUrl)
            //{
            //    string path = "";

            //    if (!browseUrl.StartsWith(referenceStrIdentity))
            //    {
            //        //it is a path
            //        path = browseUrl;
            //    }


            //    if (project.Object is VSLangProj.VSProject)
            //    {
            //        VSLangProj.VSProject vsproject = (VSLangProj.VSProject)project.Object;
            //        VSLangProj.Reference reference = null;
            //        try
            //        {
            //            reference = vsproject.References.Find(referenceStrIdentity);
            //        }
            //        catch (Exception ex)
            //        {
            //            //it failed to find one, so it must not exist. 
            //            //But it decided to error for the fun of it. :)
            //        }
            //        if (reference == null)
            //        {
            //            if (path == "")
            //                vsproject.References.Add(browseUrl);
            //            else
            //                vsproject.References.Add(path);
            //        }
            //        else
            //        {
            //            throw new Exception("Reference already exists.");
            //        }
            //    }
            //    else if (project.Object is VsWebSite.VSWebSite)
            //    {
            //        VsWebSite.VSWebSite vswebsite = (VsWebSite.VSWebSite)project.Object;
            //        VsWebSite.AssemblyReference reference = null;
            //        try
            //        {
            //            foreach (VsWebSite.AssemblyReference r in vswebsite.References)
            //            {
            //                if (r.Name == referenceStrIdentity)
            //                {
            //                    reference = r;
            //                    break;
            //                }
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //            //it failed to find one, so it must not exist. 
            //            //But it decided to error for the fun of it. :)
            //        }
            //        if (reference == null)
            //        {
            //            if (path == "")
            //                vswebsite.References.AddFromGAC(browseUrl);
            //            else
            //                vswebsite.References.AddFromFile(path);

            //        }
            //        else
            //        {
            //            throw new Exception("Reference already exists.");
            //        }
            //    }
            //    else
            //    {
            //        throw new Exception("Currently, system is only set up " +
            //                  "to do references for normal projects.");
            //    }
            //}
        }

        ////
        //// Methods In - public static class TextSelectionExtensions
        ////
        //public static int LastLine(this EnvDTE.TextSelection textSel)
        //public static int LastLine(this EnvDTE.ProjectItem currentItem)
        //public static void DisplayEntireDocument(this EnvDTE.TextSelection textSel)
        //public static void DisplayEntireDocument(this EnvDTE.ProjectItem currentItem)
    }
    public class EnvDTEBase
    {
        // begin EnvDTE fields
        public System.Type type;
        public Object obj;
        public EnvDTE.DTE dte;
        // end EnvDTE fields
        public static void CreateDirectory(string DirectoryPath)
        {
            Directory.CreateDirectory(DirectoryPath); // this creates a directory recursively
        }
        public static void CreateSolution(EnvDTEBase DTEBase, string directory, string solutionName)
        {
            // create a new solution
            CreateDirectory(directory);

            var solution = DTEBase.dte.Solution;
            solution.Create(directory, solutionName);
        }
        public static void InitializeDTEBase(EnvDTEBase DTEBase) // initialize EnvDTEBase
        {
            // begin initialize EnvDTE fields 
            DTEBase.type = System.Type.GetTypeFromProgID("VisualStudio.DTE.12.0");
            DTEBase.obj = System.Activator.CreateInstance(DTEBase.type, true);
            DTEBase.dte = (EnvDTE.DTE)DTEBase.obj; // dte is the ROOT 
            DTEBase.dte.MainWindow.Visible = true; // optional if you want to See Visual Studio doing its thing, opens visual studio and shows what it is doing
            // end initialize EnvDTE fields  
        }
        public static void SaveAndCloseDTE(EnvDTEBase DTEBase) // Saves the solution and closes the Visual Studio instance opened in CreateDTEBase(EnvDTEBase DTEBase)
        {
            // begin EnvDTE - now save and close the solution, project, items, references, settings, etc
            DTEBase.dte.ExecuteCommand("File.SaveAll");
            DTEBase.dte.Quit();
            // end EnvDTE - now save and close the solution, project, items, references, settings, etc
        }
    }
    public static class EnvDTEExtensions
    {
        public static void AddFromFile(this EnvDTE.Project project, List<string> path, string file)
        {
            ProjectItems pi = project.ProjectItems;
            for (int i = 0; i < path.Count; i++)
            {
                pi = pi.Item(path[i]).ProjectItems;
            }
            pi.AddFromFile(file);
        }
        public static EnvDTE.ProjectItem AddFolders(this EnvDTE.Project project, List<string> path) // adds the directory(s) recursively then returns a ProjectItem
        { // and yes I know it is directories
            ProjectItem pItem = null;
            ProjectItems pItems = project.ProjectItems;
            for (int i = 0; i < path.Count; i++)
            {
                if (CheckIfFileExistsInProject(pItems, path[i]))
                { // go there if exists
                    pItems = pItems.Item(path[i]).ProjectItems;
                }
                else
                { // add if it doesn't
                    pItem = pItems.AddFolder(path[i]);
                    pItems = pItems.Item(path[i]).ProjectItems;
                }
            }
            return pItem;
        }
        private static bool CheckIfFileExistsInProject(ProjectItems projectItems, string itemName)
        {
            // initial value
            bool fileExists = false;
            // iterate project items
            foreach (ProjectItem projectItem in projectItems)
            {
                // if the name matches
                if (projectItem.Name == itemName)
                {
                    // abort this add, file already exists
                    fileExists = true;
                    // break out of loop
                    break;
                }
                else if ((projectItem.ProjectItems != null) && (projectItem.ProjectItems.Count > 0))
                {
                    // check if the file exists in the project
                    fileExists = CheckIfFileExistsInProject(projectItem.ProjectItems, itemName);
                    // if the file does exist
                    if (fileExists)
                    {
                        // abort this add, file already exists
                        fileExists = true;
                        // break out of loop
                        break;
                    }
                }
            }
            // return value
            return fileExists;
        }
        private static bool CheckIfFileExistsInProject(Project project, string itemName)
        {
            // initial value
            bool fileExists = false;
            // iterate project items
            foreach (ProjectItem projectItem in project.ProjectItems)
            {
                // if the name matches
                if (projectItem.Name == itemName)
                {
                    // abort this add, file already exists
                    fileExists = true;
                    // break out of loop
                    break;
                }
                else if ((projectItem.ProjectItems != null) && (projectItem.ProjectItems.Count > 0))
                {
                    // check if the file exists in the project
                    fileExists = CheckIfFileExistsInProject(projectItem.ProjectItems, itemName);
                    // if the file does exist
                    if (fileExists)
                    {
                        // abort this add, file already exists
                        fileExists = true;
                        // break out of loop
                        break;
                    }
                }
            }
            // return value
            return fileExists;
        }
        public static void ExamineCodeElement(CodeElement codeElement, int tabs) // recursively examine code elements
        {
            tabs++;
            try
            {
                Console.WriteLine(new string('\t', tabs) + "{0} {1}", codeElement.Name, codeElement.Kind.ToString());
                foreach (CodeElement childElement in codeElement.Children)
                {
                    ExamineCodeElement(childElement, tabs);
                }
            }
            catch
            {
                Console.WriteLine(new string('\t', tabs) + "codeElement without name: {0}", codeElement.Kind.ToString());
            }
        }
        public static EnvDTE.ProjectItem FindProjectItemInProject(Project project, string name, bool recursive) // directory tree recursive find of files
        {
            ProjectItem projectItem = null;
            if (project.Kind != EnvDTEConstants.vsProjectKindSolutionItems)
            {
                if (project.ProjectItems != null && project.ProjectItems.Count > 0)
                {
                    projectItem = FindItemByName(project.ProjectItems, name, recursive);
                }
            }
            else
            {
                // if solution folder, one of its ProjectItems might be a real project
                foreach (ProjectItem item in project.ProjectItems)
                {
                    Project realProject = item.Object as Project;

                    if (realProject != null)
                    {
                        projectItem = EnvDTEExtensions.FindProjectItemInProject(realProject, name, recursive);

                        if (projectItem != null)
                        {
                            break;
                        }
                    }
                }
            }
            return projectItem;
        }
        public static EnvDTE.ProjectItem FindItemByName(ProjectItems collection, string name, bool recursive)
        {
            if (collection != null)
            {
                foreach (ProjectItem item1 in collection)
                {
                    if (item1.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return item1;
                    }
                    if (recursive)
                    {
                        ProjectItem item2 = EnvDTEExtensions.FindItemByName(item1.ProjectItems, name, recursive);
                        if (item2 != null)
                        {
                            return item2;
                        }
                    }
                }
            }

            return null;
        }
        public static EnvDTE.ProjectItem GetItemByName(this EnvDTE.Project project, string itemName) // requires file .ext (.cs) to be added to the creation name
        {
            // TODO make this call FindItemByName(ProjectItems collection, string name, bool recursive)
            EnvDTE.ProjectItem returnItem = null;
            foreach (EnvDTE.ProjectItem item in project.ProjectItems)
            {
                if (item.Name == itemName)
                {
                    returnItem = item;
                    break;
                }
            }
            return returnItem;
        }
        public static EnvDTE.ProjectItem GetItemByName(this EnvDTE.ProjectItems pItems, string itemName) // requires file .ext (.cs) to be added to the creation name
        {
            // TODO make this call FindItemByName(ProjectItems collection, string name, bool recursive)
            EnvDTE.ProjectItem returnItem = null;
            foreach (EnvDTE.ProjectItem item in pItems)
            {
                if (item.Name == itemName)
                {
                    returnItem = item;
                    break;
                }
            }
            return returnItem;
        }
        public static EnvDTE.Project GetProjectByName(this EnvDTE.Projects projects, string projectName)
        {
            EnvDTE.Project returnProject = null;
            foreach (EnvDTE.Project proj in projects)
            {
                if (proj.Name == projectName)
                {
                    returnProject = proj;
                    break;
                }
            }
            return returnProject;
        }

        // TODO create - public static EnvDTE.Solution GetSolutionByName(this EnvDTE.Solution sol, string Name) { }

        public static void Reformat(this EnvDTE.ProjectItem projItem) // Reformats the code the same way ctrl-E-D/K-D do 
        {
            if (projItem != null)
            {
                if (projItem.Kind == EnvDTEConstants.vsDocumentKindText || projItem.Kind == EnvDTEConstants.vsProjectItemKindPhysicalFile)
                {
                    try // redo this to check if it is a directory before Reformat - fixed with the above two if statements
                    {
                        TextSelection objSel = projItem.Document.Selection;  // this fails for directories
                        objSel.SelectAll();
                        objSel.SmartFormat();
                        objSel.GotoLine(1, false); // (lineNum, select bool)
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Tried to Reformat() something (ie: dir) other than a code file, \nException message : " + ex.Message);
                    }
                }
            }
        }
        public static void SaveAndClose(this EnvDTE.ProjectItem projItem) // this saves and closes the open edit windows in VS, USE IT after each file is completed!!! or code gen will slow down a lot. 
        {
            if (projItem != null)
            {
                if (projItem.Kind == EnvDTEConstants.vsDocumentKindText || projItem.Kind == EnvDTEConstants.vsProjectItemKindPhysicalFile)
                {
                    try // redo this to check it is a directory before Reformat - fixed with the above two if statements
                    {
                        projItem.Reformat(); // clean up the formating
                        projItem.Document.Close(vsSaveChanges.vsSaveChangesYes); // save the changes
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("SaveAndClose() - Tried to Reformat() or save and Close something (ie: dir) other than a code file, \nException message : " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("SaveAndClose() - projItem = {0}", projItem.Kind.ToString());
                }
            }
            else
            {
                Console.WriteLine("SaveAndClose() - attempted to save with, projItem = null");
            }
        }
    }
    internal class EnvDTEConstants // without this the Constants for EnvDTE will flip an error at build/runtime
    {
        public const string vsAddInCmdGroup = "{1E58696E-C90F-11D2-AAB2-00C04F688DDE}";
        public const string vsCATIDDocument = "{610d4611-d0d5-11d2-8599-006097c68e81}";
        public const string vsCATIDGenericProject = "{610d4616-d0d5-11d2-8599-006097c68e81}";
        public const string vsCATIDMiscFilesProject = "{610d4612-d0d5-11d2-8599-006097c68e81}";
        public const string vsCATIDMiscFilesProjectItem = "{610d4613-d0d5-11d2-8599-006097c68e81}";
        public const string vsCATIDSolution = "{52AEFF70-BBD8-11d2-8598-006097C68E81}";
        public const string vsCATIDSolutionBrowseObject = "{A2392464-7C22-11d3-BDCA-00C04F688E50}";
        public const string vsContextDebugging = "{ADFC4E61-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextDesignMode = "{ADFC4E63-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextEmptySolution = "{ADFC4E65-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextFullScreenMode = "{ADFC4E62-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextMacroRecording = "{04BBF6A5-4697-11D2-890E-0060083196C6}";
        public const string vsContextMacroRecordingToolbar = "{85A70471-270A-11D2-88F9-0060083196C6}";
        public const string vsContextNoSolution = "{ADFC4E64-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextSolutionBuilding = "{ADFC4E60-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextSolutionHasMultipleProjects = "{93694FA0-0397-11D1-9F4E-00A0C911004F}";
        public const string vsContextSolutionHasSingleProject = "{ADFC4E66-0397-11D1-9F4E-00A0C911004F}";
        public const string vsDocumentKindBinary = "{25834150-CD7E-11D0-92DF-00A0C9138C45}";
        public const string vsDocumentKindHTML = "{C76D83F8-A489-11D0-8195-00A0C91BBEE3}";
        public const string vsDocumentKindResource = "{00000000-0000-0000-0000-000000000000}";
        public const string vsDocumentKindText = "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}";      // text document 
        public const string vsext_GUID_AddItemWizard = "{0F90E1D1-4999-11D1-B6D1-00A0C90F2744}";
        public const string vsext_GUID_NewProjectWizard = "{0F90E1D0-4999-11D1-B6D1-00A0C90F2744}";
        public const string vsext_vk_Code = "{7651A701-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsext_vk_Debugging = "{7651A700-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsext_vk_Designer = "{7651A702-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsext_vk_Primary = "{00000000-0000-0000-0000-000000000000}";
        public const string vsext_vk_TextView = "{7651A703-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsext_wk_AutoLocalsWindow = "{F2E84780-2AF1-11D1-A7FA-00A0C9110051}";
        public const string vsext_wk_CallStackWindow = "{0504FF91-9D61-11D0-A794-00A0C9110051}";
        public const string vsext_wk_ClassView = "{C9C0AE26-AA77-11D2-B3F0-0000F87570EE}";
        public const string vsext_wk_ContextWindow = "{66DBA47C-61DF-11D2-AA79-00C04F990343}";
        public const string vsext_wk_ImmedWindow = "{98731960-965C-11D0-A78F-00A0C9110051}";
        public const string vsext_wk_LocalsWindow = "{4A18F9D0-B838-11D0-93EB-00A0C90F2734}";
        public const string vsext_wk_ObjectBrowser = "{269A02DC-6AF8-11D3-BDC4-00C04F688E50}";
        public const string vsext_wk_OutputWindow = "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}";
        public const string vsext_wk_PropertyBrowser = "{EEFA5220-E298-11D0-8F78-00A0C9110057}";
        public const string vsext_wk_SProjectWindow = "{3AE79031-E1BC-11D0-8F78-00A0C9110057}";
        public const string vsext_wk_TaskList = "{4A9B7E51-AA16-11D0-A8C5-00A0C921A4D2}";
        public const string vsext_wk_ThreadWindow = "{E62CE6A0-B439-11D0-A79D-00A0C9110051}";
        public const string vsext_wk_Toolbox = "{B1E99781-AB81-11D0-B683-00AA00A3EE26}";
        public const string vsext_wk_WatchWindow = "{90243340-BD7A-11D0-93EF-00A0C90F2734}";
        public const string vsMiscFilesProjectUniqueName = "<MiscFiles>";
        public const string vsProjectItemKindMisc = "{66A2671F-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectItemKindPhysicalFile = "{6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}";
        public const string vsProjectItemKindPhysicalFolder = "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}";
        public const string vsProjectItemKindSolutionItems = "{66A26722-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectItemKindSubProject = "{EA6618E8-6E24-4528-94BE-6889FE16485C}";
        public const string vsProjectItemKindVirtualFolder = "{6BB5F8F0-4483-11D3-8BCF-00C04F8EC28C}";
        public const string vsProjectItemsKindMisc = "{66A2671E-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectItemsKindSolutionItems = "{66A26721-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectKindMisc = "{66A2671D-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectKindSolutionItems = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        public const string vsProjectKindUnmodeled = "{67294A52-A4F0-11D2-AA88-00C04F688DDE}";
        public const string vsProjectsKindSolution = "{96410B9F-3542-4A14-877F-BC7227B51D3B}";
        public const string vsSolutionItemsProjectUniqueName = "<SolnItems>";
        public const string vsViewKindAny = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}";
        public const string vsViewKindCode = "{7651A701-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsViewKindDebugging = "{7651A700-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsViewKindDesigner = "{7651A702-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsViewKindPrimary = "{00000000-0000-0000-0000-000000000000}";
        public const string vsViewKindTextView = "{7651A703-06E5-11D1-8EBD-00A0C90F26EA}";
        public const string vsWindowKindAutoLocals = "{F2E84780-2AF1-11D1-A7FA-00A0C9110051}";
        public const string vsWindowKindCallStack = "{0504FF91-9D61-11D0-A794-00A0C9110051}";
        public const string vsWindowKindClassView = "{C9C0AE26-AA77-11D2-B3F0-0000F87570EE}";
        public const string vsWindowKindCommandWindow = "{28836128-FC2C-11D2-A433-00C04F72D18A}";
        public const string vsWindowKindDocumentOutline = "{25F7E850-FFA1-11D0-B63F-00A0C922E851}";
        public const string vsWindowKindDynamicHelp = "{66DBA47C-61DF-11D2-AA79-00C04F990343}";
        public const string vsWindowKindFindReplace = "{CF2DDC32-8CAD-11D2-9302-005345000000}";
        public const string vsWindowKindFindResults1 = "{0F887920-C2B6-11D2-9375-0080C747D9A0}";
        public const string vsWindowKindFindResults2 = "{0F887921-C2B6-11D2-9375-0080C747D9A0}";
        public const string vsWindowKindFindSymbol = "{53024D34-0EF5-11D3-87E0-00C04F7971A5}";
        public const string vsWindowKindFindSymbolResults = "{68487888-204A-11D3-87EB-00C04F7971A5}";
        public const string vsWindowKindLinkedWindowFrame = "{9DDABE99-1D02-11D3-89A1-00C04F688DDE}";
        public const string vsWindowKindLocals = "{4A18F9D0-B838-11D0-93EB-00A0C90F2734}";
        public const string vsWindowKindMacroExplorer = "{07CD18B4-3BA1-11D2-890A-0060083196C6}";
        public const string vsWindowKindMainWindow = "{9DDABE98-1D02-11D3-89A1-00C04F688DDE}";
        public const string vsWindowKindObjectBrowser = "{269A02DC-6AF8-11D3-BDC4-00C04F688E50}";
        public const string vsWindowKindOutput = "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}";
        public const string vsWindowKindProperties = "{EEFA5220-E298-11D0-8F78-00A0C9110057}";
        public const string vsWindowKindResourceView = "{2D7728C2-DE0A-45b5-99AA-89B609DFDE73}";
        public const string vsWindowKindServerExplorer = "{74946827-37A0-11D2-A273-00C04F8EF4FF}";
        public const string vsWindowKindSolutionExplorer = "{3AE79031-E1BC-11D0-8F78-00A0C9110057}";
        public const string vsWindowKindTaskList = "{4A9B7E51-AA16-11D0-A8C5-00A0C921A4D2}";
        public const string vsWindowKindThread = "{E62CE6A0-B439-11D0-A79D-00A0C9110051}";
        public const string vsWindowKindToolbox = "{B1E99781-AB81-11D0-B683-00AA00A3EE26}";
        public const string vsWindowKindWatch = "{90243340-BD7A-11D0-93EF-00A0C90F2734}";
        public const string vsWindowKindWebBrowser = "{E8B06F52-6D01-11D2-AA7D-00C04F990343}";
        public const string vsWizardAddItem = "{0F90E1D1-4999-11D1-B6D1-00A0C90F2744}";
        public const string vsWizardAddSubProject = "{0F90E1D2-4999-11D1-B6D1-00A0C90F2744}";
        public const string vsWizardNewProject = "{0F90E1D0-4999-11D1-B6D1-00A0C90F2744}";
    }
    public static class FileSystemMethods
    {
        public static void ConvertDirectoryStringToList(this List<string> listStr, string dirStr)
        {
            string[] directories = dirStr.Split(Path.DirectorySeparatorChar); // adds empty trailing dir name due to final slash "\"
            //listStr = new List<string>(directories.ToList<string>()); // only valid in this context, will not return a value
            for (int i = 0; i < directories.Length; i++)
            {
                if (directories[i].Trim() != String.Empty) // empty entry - trailing and leading slash removal
                {
                    listStr.Add(directories[i]);
                }
            }
        }
        public static void DeleteFilesAndDirectories(string targetDirectory)
        {
            Console.WriteLine("\n\nDeleting all files in {0}\n", targetDirectory);
            DirectoryInfo dirInfo = new DirectoryInfo(targetDirectory);
            RecursiveDelete(dirInfo);
        }
        public static void RecursiveDelete(DirectoryInfo baseDir)
        {
            if (!baseDir.Exists)
                return;

            foreach (var dir in baseDir.EnumerateDirectories())
            {
                RecursiveDelete(dir);
            }
            if (baseDir.Exists)
            {
                baseDir.Delete(true); // fails some times, probably the google drive sync locking the directory up, RERUN UNTIL IT WORKS
            }
        }
    }
    public static class TextSelectionExtensions
    {
        //
        // https://msdn.microsoft.com/en-us/library/envdte.textselection.aspx 
        // some of properties that do not show up in intellisense are listed there
        // (ie CurrentLine, )
        //

        public static void DisplayEntireDocument(this EnvDTE.TextSelection textSel)
        {
            int lastLine = LastLine(textSel);
            for (int i = 1; i < lastLine; i++)
            {
                textSel.GotoLine(i, true);
                Console.WriteLine(textSel.Text);
            }
        }
        public static void DisplayEntireDocument(this EnvDTE.ProjectItem currentItem)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            DisplayEntireDocument(textSel);
        }
        public static int GetLineNumberContaining(this EnvDTE.TextSelection textSel, string searchString)
        {
            int lastLine = LastLine(textSel);
            for (int i = 1; i < lastLine; i++)
            {
                textSel.GotoLine(i, true);
                if (textSel.Text.Contains(searchString))
                {
                    return i; // line number
                }
            }
            return -1; // failed
        }
        public static int GetLineNumberContaining(this EnvDTE.ProjectItem currentItem, string searchString)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            return textSel.GetLineNumberContaining(searchString);
        }
        public static TextSelection GetTextSelectionFromProjectItem(this EnvDTE.ProjectItem currentItem)
        {
            return currentItem.Document.Selection;
        }
        public static void InsertAboveThisText(this EnvDTE.TextSelection textSel, string searchString, string insertString)
        {
            // search from first to last line looking for searchString
            int currentLineNumber = textSel.GetLineNumberContaining(searchString);
            textSel.GotoLine(currentLineNumber);
            textSel.StartOfLine();
            textSel.Insert("\n");
            textSel.GotoLine(currentLineNumber);
            textSel.Insert(insertString);
        }
        public static void InsertAboveThisText(this EnvDTE.ProjectItem currentItem, string searchString, string insertString)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            InsertAboveThisText(textSel, searchString, insertString);
        }
        public static void InsertAboveThisText(this EnvDTE.TextSelection textSel, string searchString, List<string> listOfStrings)
        {
            // search from first to last line looking for searchString
            int currentLineNumber = textSel.GetLineNumberContaining(searchString);
            textSel.GotoLine(currentLineNumber);
            textSel.StartOfLine();
            textSel.Insert("\n");
            textSel.GotoLine(currentLineNumber);
            foreach (string str in listOfStrings)
            {
                textSel.Insert(str);
                textSel.Insert("\n");
            }

        }
        public static void InsertAboveThisText(this EnvDTE.ProjectItem currentItem, string searchString, List<string> listOfStrings)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            InsertAboveThisText(textSel, searchString, listOfStrings);
        }
        public static void InsertBelowThisText(this EnvDTE.TextSelection textSel, string searchString, string insertString)
        {
            // search from first to last line looking for searchString
            int currentLineNumber = textSel.GetLineNumberContaining(searchString);
            textSel.GotoLine(currentLineNumber);
            textSel.EndOfLine();
            textSel.Insert("\n");
            textSel.GotoLine(currentLineNumber + 1);
            textSel.Insert(insertString);
        }
        public static void InsertBelowThisText(this EnvDTE.ProjectItem currentItem, string searchString, string insertString)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            InsertBelowThisText(textSel, searchString, insertString);
        }
        public static void InsertBelowThisText(this EnvDTE.TextSelection textSel, string searchString, List<string> listOfStrings)
        {
            // search from first to last line looking for searchString
            int currentLineNumber = textSel.GetLineNumberContaining(searchString);
            textSel.GotoLine(currentLineNumber);
            textSel.EndOfLine();
            textSel.Insert("\n");
            textSel.GotoLine(currentLineNumber + 1);
            foreach (string str in listOfStrings)
            {
                textSel.Insert(str);
                textSel.Insert("\n");
            }
        }
        public static void InsertBelowThisText(this EnvDTE.ProjectItem currentItem, string searchString, List<string> listOfStrings)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            InsertBelowThisText(textSel, searchString, listOfStrings);
        }
        public static int LastLine(this EnvDTE.TextSelection textSel)
        {
            textSel.EndOfDocument(true);
            return textSel.CurrentLine;
        }
        public static int LastLine(this EnvDTE.ProjectItem currentItem)
        {
            TextSelection textSel = GetTextSelectionFromProjectItem(currentItem);
            return LastLine(textSel);
        }

        // TODO ReplaceText(this EnvDTE.ProjectItem currentItem, string searchText, string replacementText) {}
        // TODO ReplaceToken(this EnvDTE.ProjectItem currentItem, string searchText, string replacementText) {}

    }
}
