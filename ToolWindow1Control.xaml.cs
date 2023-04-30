namespace lab3
{
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;
    using EnvDTE;
    using EnvDTE80;
    using System.Text.RegularExpressions;
    using System.Collections.Generic;
    using Microsoft.VisualStudio.Shell;
    using System;

    public class StatisticsSet
    {
        public string FunctionName { get; set; }
        public string LinesCount { get; set; }
        public string PureLinesCount { get; set; }
        public string KeywordCount { get; set; }
    }

    /// <summary>
    /// Interaction logic for ToolWindow1Control.
    /// </summary>
    public partial class ToolWindow1Control : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow1Control"/> class.
        /// </summary>

        private List<StatisticsSet> StatisticsTable;

        public ToolWindow1Control()
        {
            this.InitializeComponent();

            StatisticsTable = new List<StatisticsSet>();

            StatisticsList.ItemsSource = StatisticsTable;
        }

        private List<CodeFunction> FunctionList;

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]

        private void ParseFunction(int i)
        {
            Dispatcher.VerifyAccess();

            TextPoint StartPoint = FunctionList[i].GetStartPoint(vsCMPart.vsCMPartHeader);

            TextPoint FinishPoint;

            string OriginalContent;

            if (i != FunctionList.Count - 1)
            {
                FinishPoint = FunctionList[i + 1].GetStartPoint(vsCMPart.vsCMPartHeader);

                OriginalContent = StartPoint.CreateEditPoint().GetText(FinishPoint);
            }
            else
            {
                FinishPoint = FunctionList[i].GetEndPoint().CreateEditPoint();

                EditPoint EP = FunctionList[i].GetEndPoint().CreateEditPoint();

                EP.EndOfDocument();

                OriginalContent = StartPoint.CreateEditPoint().GetText(EP);
            }

            string FunctionContent = OriginalContent;

            string PureContent = "";

            Match quote;

            Match apostrophe;

            Match comment;

            Match scomment;

            Match sscomment;

            int QuoteIdx;

            int ApostropheIdx;

            int CommentIdx;

            int ScommentIdx;

            int SscommentIdx;

            string temp;

            do
            {
                quote = Regex.Match(FunctionContent, @"("")");

                apostrophe = Regex.Match(FunctionContent, @"(')");

                comment = Regex.Match(FunctionContent, @"(/\*([^*]|[\r\n]|(\*+([^*/]|[\r\n])))*\*+/)");

                scomment = Regex.Match(FunctionContent, @"(//.*\\((\r\n)|(\r)|(\n)))");

                sscomment = Regex.Match(FunctionContent, @"(//.*((\r\n)|(\r)|(\n)))");

                QuoteIdx = int.MaxValue;

                ApostropheIdx = int.MaxValue;

                CommentIdx = int.MaxValue;

                ScommentIdx = int.MaxValue;

                SscommentIdx = int.MaxValue;

                if ("" != quote.Value)
                {
                    QuoteIdx = quote.Index;
                }

                if ("" != apostrophe.Value)
                {
                    ApostropheIdx = apostrophe.Index;
                }

                if ("" != comment.Value)
                {
                    CommentIdx = comment.Index;
                }

                if ("" != scomment.Value)
                {
                    ScommentIdx = scomment.Index;
                }

                if ("" != sscomment.Value)
                {
                    SscommentIdx = sscomment.Index;
                }

                if (QuoteIdx < ApostropheIdx && QuoteIdx < CommentIdx && QuoteIdx < ScommentIdx && QuoteIdx < SscommentIdx)
                {
                    temp = FunctionContent;

                    temp = temp.Remove(QuoteIdx);

                    PureContent += temp + " 0";

                    FunctionContent = FunctionContent.Substring(QuoteIdx + quote.Length);

                    Match NewQuote;

                    Match NewBackQuote;

                    Match NewNewLine;

                    Match NewNewBackLine;

                    int NewQuoteIdx;

                    int NewBackQuoteIdx;

                    int NewNewLineIdx;

                    int NewNewBackLineIdx;

                    while (true)
                    {
                        NewQuote = Regex.Match(FunctionContent, @"("")");

                        NewBackQuote = Regex.Match(FunctionContent, @"((\\)+"")");

                        NewNewLine = Regex.Match(FunctionContent, @"((\r\n)|(\r)|(\n))");

                        NewNewBackLine = Regex.Match(FunctionContent, @"((\\)+((\r\n)|(\r)|(\n)))");

                        NewQuoteIdx = int.MaxValue;

                        NewBackQuoteIdx = int.MaxValue;

                        NewNewLineIdx = int.MaxValue;

                        NewNewBackLineIdx = int.MaxValue;

                        if ("" != NewQuote.Value)
                        {
                            NewQuoteIdx = NewQuote.Index;
                        }

                        if ("" != NewBackQuote.Value)
                        {
                            NewBackQuoteIdx = NewBackQuote.Index;
                        }

                        if ("" != NewNewLine.Value)
                        {
                            NewNewLineIdx = NewNewLine.Index;
                        }

                        if ("" != NewNewBackLine.Value)
                        {
                            NewNewBackLineIdx = NewNewBackLine.Index;
                        }

                        if (NewQuoteIdx < NewBackQuoteIdx && NewQuoteIdx < NewNewLineIdx && NewQuoteIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0";

                            FunctionContent = FunctionContent.Substring(NewQuoteIdx + NewQuote.Length);

                            break;
                        }
                        else if (NewBackQuoteIdx < NewQuoteIdx && NewBackQuoteIdx < NewNewLineIdx && NewBackQuoteIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0";

                            FunctionContent = FunctionContent.Substring(NewBackQuoteIdx + NewBackQuote.Length);

                            if (NewBackQuote.Length % 2 != 0)
                            {
                                break;
                            }
                        }
                        else if (NewNewLineIdx < NewQuoteIdx && NewNewLineIdx < NewBackQuoteIdx && NewNewLineIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0\r\n";

                            FunctionContent = FunctionContent.Substring(NewNewLineIdx + NewNewLine.Length);

                            break;
                        }
                        else if (NewNewBackLineIdx < NewQuoteIdx && NewNewBackLineIdx < NewBackQuoteIdx && NewNewBackLineIdx < NewNewLineIdx)
                        {
                            PureContent += " 0\r\n";

                            FunctionContent = FunctionContent.Substring(NewNewBackLineIdx + NewNewBackLine.Length);

                            if (NewNewBackLine.Length % 2 == 0)
                            {
                                break;
                            }

                            do
                            {
                                NewQuote = Regex.Match(FunctionContent, @"("")");

                                NewBackQuote = Regex.Match(FunctionContent, @"((\\)+"")");

                                NewNewLine = Regex.Match(FunctionContent, @"((\r\n)|(\r)|(\n))");

                                NewNewBackLine = Regex.Match(FunctionContent, @"((\\)+((\r\n)|(\r)|(\n)))");

                                NewQuoteIdx = int.MaxValue;

                                NewBackQuoteIdx = int.MaxValue;

                                NewNewLineIdx = int.MaxValue;

                                NewNewBackLineIdx = int.MaxValue;

                                if ("" != NewQuote.Value)
                                {
                                    NewQuoteIdx = NewQuote.Index;
                                }

                                if ("" != NewBackQuote.Value)
                                {
                                    NewBackQuoteIdx = NewBackQuote.Index;
                                }

                                if ("" != NewNewLine.Value)
                                {
                                    NewNewLineIdx = NewNewLine.Index;
                                }

                                if ("" != NewNewBackLine.Value)
                                {
                                    NewNewBackLineIdx = NewNewBackLine.Index;
                                }

                                if (NewQuoteIdx < NewBackQuoteIdx && NewQuoteIdx < NewNewLineIdx && NewQuoteIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0";

                                    FunctionContent = FunctionContent.Substring(NewQuoteIdx + NewQuote.Length);

                                    break;
                                }
                                else if (NewBackQuoteIdx < NewQuoteIdx && NewBackQuoteIdx < NewNewLineIdx && NewBackQuoteIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0";

                                    FunctionContent = FunctionContent.Substring(NewBackQuoteIdx + NewBackQuote.Length);

                                    if (NewBackQuote.Length % 2 != 0)
                                    {
                                        break;
                                    }
                                }
                                else if (NewNewLineIdx < NewQuoteIdx && NewNewLineIdx < NewBackQuoteIdx && NewNewLineIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0\r\n";

                                    FunctionContent = FunctionContent.Substring(NewNewLineIdx + NewNewLine.Length);

                                    break;
                                }
                                else if (NewNewBackLineIdx < NewQuoteIdx && NewNewBackLineIdx < NewBackQuoteIdx && NewNewBackLineIdx < NewNewLineIdx)
                                {
                                    PureContent += " 0\r\n";

                                    FunctionContent = FunctionContent.Substring(NewNewBackLineIdx + NewNewBackLine.Length);

                                    if (NewNewBackLine.Length % 2 == 0)
                                    {
                                        break;
                                    }
                                }

                            } while (int.MaxValue != NewQuoteIdx || int.MaxValue != NewBackQuoteIdx || int.MaxValue != NewNewLineIdx || int.MaxValue != NewNewBackLineIdx);

                            break;
                        }
                    }
                }
                else if (ApostropheIdx < QuoteIdx && ApostropheIdx < CommentIdx && ApostropheIdx < ScommentIdx && ApostropheIdx < SscommentIdx)
                {
                    temp = FunctionContent;

                    temp = temp.Remove(ApostropheIdx);

                    PureContent += temp + " 0";

                    FunctionContent = FunctionContent.Substring(ApostropheIdx + apostrophe.Length);

                    Match NewQuote;

                    Match NewBackQuote;

                    Match NewNewLine;

                    Match NewNewBackLine;

                    int NewQuoteIdx;

                    int NewBackQuoteIdx;

                    int NewNewLineIdx;

                    int NewNewBackLineIdx;

                    while (true)
                    {
                        NewQuote = Regex.Match(FunctionContent, @"(')");

                        NewBackQuote = Regex.Match(FunctionContent, @"((\\)+')");

                        NewNewLine = Regex.Match(FunctionContent, @"((\r\n)|(\r)|(\n))");

                        NewNewBackLine = Regex.Match(FunctionContent, @"((\\)+((\r\n)|(\r)|(\n)))");

                        NewQuoteIdx = int.MaxValue;

                        NewBackQuoteIdx = int.MaxValue;

                        NewNewLineIdx = int.MaxValue;

                        NewNewBackLineIdx = int.MaxValue;

                        if ("" != NewQuote.Value)
                        {
                            NewQuoteIdx = NewQuote.Index;
                        }

                        if ("" != NewBackQuote.Value)
                        {
                            NewBackQuoteIdx = NewBackQuote.Index;
                        }

                        if ("" != NewNewLine.Value)
                        {
                            NewNewLineIdx = NewNewLine.Index;
                        }

                        if ("" != NewNewBackLine.Value)
                        {
                            NewNewBackLineIdx = NewNewBackLine.Index;
                        }

                        if (NewQuoteIdx < NewBackQuoteIdx && NewQuoteIdx < NewNewLineIdx && NewQuoteIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0";

                            FunctionContent = FunctionContent.Substring(NewQuoteIdx + NewQuote.Length);

                            break;
                        }
                        else if (NewBackQuoteIdx < NewQuoteIdx && NewBackQuoteIdx < NewNewLineIdx && NewBackQuoteIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0";

                            FunctionContent = FunctionContent.Substring(NewBackQuoteIdx + NewBackQuote.Length);

                            if (NewBackQuote.Length % 2 != 0)
                            {
                                break;
                            }
                        }
                        else if (NewNewLineIdx < NewQuoteIdx && NewNewLineIdx < NewBackQuoteIdx && NewNewLineIdx < NewNewBackLineIdx)
                        {
                            PureContent += " 0\r\n";

                            FunctionContent = FunctionContent.Substring(NewNewLineIdx + NewNewLine.Length);

                            break;
                        }
                        else if (NewNewBackLineIdx < NewQuoteIdx && NewNewBackLineIdx < NewBackQuoteIdx && NewNewBackLineIdx < NewNewLineIdx)
                        {
                            PureContent += " 0\r\n";

                            FunctionContent = FunctionContent.Substring(NewNewBackLineIdx + NewNewBackLine.Length);

                            if (NewNewBackLine.Length % 2 == 0)
                            {
                                break;
                            }

                            do
                            {
                                NewQuote = Regex.Match(FunctionContent, @"(')");

                                NewBackQuote = Regex.Match(FunctionContent, @"((\\)+')");

                                NewNewLine = Regex.Match(FunctionContent, @"((\r\n)|(\r)|(\n))");

                                NewNewBackLine = Regex.Match(FunctionContent, @"((\\)+((\r\n)|(\r)|(\n)))");

                                NewQuoteIdx = int.MaxValue;

                                NewBackQuoteIdx = int.MaxValue;

                                NewNewLineIdx = int.MaxValue;

                                NewNewBackLineIdx = int.MaxValue;

                                if ("" != NewQuote.Value)
                                {
                                    NewQuoteIdx = NewQuote.Index;
                                }

                                if ("" != NewBackQuote.Value)
                                {
                                    NewBackQuoteIdx = NewBackQuote.Index;
                                }

                                if ("" != NewNewLine.Value)
                                {
                                    NewNewLineIdx = NewNewLine.Index;
                                }

                                if ("" != NewNewBackLine.Value)
                                {
                                    NewNewBackLineIdx = NewNewBackLine.Index;
                                }

                                if (NewQuoteIdx < NewBackQuoteIdx && NewQuoteIdx < NewNewLineIdx && NewQuoteIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0";

                                    FunctionContent = FunctionContent.Substring(NewQuoteIdx + NewQuote.Length);

                                    break;
                                }
                                else if (NewBackQuoteIdx < NewQuoteIdx && NewBackQuoteIdx < NewNewLineIdx && NewBackQuoteIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0";

                                    FunctionContent = FunctionContent.Substring(NewBackQuoteIdx + NewBackQuote.Length);

                                    if (NewBackQuote.Length % 2 != 0)
                                    {
                                        break;
                                    }
                                }
                                else if (NewNewLineIdx < NewQuoteIdx && NewNewLineIdx < NewBackQuoteIdx && NewNewLineIdx < NewNewBackLineIdx)
                                {
                                    PureContent += " 0\r\n";

                                    FunctionContent = FunctionContent.Substring(NewNewLineIdx + NewNewLine.Length);

                                    break;
                                }
                                else if (NewNewBackLineIdx < NewQuoteIdx && NewNewBackLineIdx < NewBackQuoteIdx && NewNewBackLineIdx < NewNewLineIdx)
                                {
                                    PureContent += " 0\r\n";

                                    FunctionContent = FunctionContent.Substring(NewNewBackLineIdx + NewNewBackLine.Length);

                                    if (NewNewBackLine.Length % 2 == 0)
                                    {
                                        break;
                                    }
                                }

                            } while (int.MaxValue != NewQuoteIdx || int.MaxValue != NewBackQuoteIdx || int.MaxValue != NewNewLineIdx || int.MaxValue != NewNewBackLineIdx);
                        }
                    }
                }
                else if (CommentIdx < QuoteIdx && CommentIdx < ApostropheIdx && CommentIdx < ScommentIdx && CommentIdx < SscommentIdx)
                {
                    temp = FunctionContent;

                    temp = temp.Remove(CommentIdx);

                    PureContent += temp;

                    temp = FunctionContent;

                    temp = temp.Remove(CommentIdx + comment.Length);

                    temp = temp.Substring(CommentIdx);

                    int nn = Regex.Matches(temp, @"((\r\n)|(\r)|(\n))").Count;

                    while (0 != nn)
                    {
                        PureContent = "\r\n" + PureContent;

                        nn--;
                    }

                    FunctionContent = FunctionContent.Substring(CommentIdx + comment.Length);
                }
                else if (ScommentIdx < QuoteIdx && ScommentIdx < ApostropheIdx && ScommentIdx < CommentIdx && ScommentIdx <= SscommentIdx)
                {
                    temp = FunctionContent;

                    temp = temp.Remove(ScommentIdx);

                    PureContent += temp;

                    do
                    {
                        PureContent += "\r\n";

                        FunctionContent = FunctionContent.Substring(ScommentIdx + scomment.Length);

                        scomment = Regex.Match(FunctionContent, @"(.*\\((\r\n)|(\r)|(\n)))");

                    } while (0 == (ScommentIdx = scomment.Index) && "" != scomment.Value);

                    PureContent += "\r\n";

                    FunctionContent = FunctionContent.Substring(Regex.Match(FunctionContent, @"(.*((\r\n)|(\r)|(\n)))").Index + Regex.Match(FunctionContent, @"(.*((\r\n)|(\r)|(\n)))").Length);

                }
                else if (SscommentIdx < QuoteIdx && SscommentIdx < ApostropheIdx && SscommentIdx < CommentIdx && SscommentIdx < ScommentIdx)
                {
                    temp = FunctionContent;

                    temp = temp.Remove(SscommentIdx);

                    PureContent += temp + "\r\n";

                    FunctionContent = FunctionContent.Substring(SscommentIdx + sscomment.Length);
                }

            } while (int.MaxValue != QuoteIdx || int.MaxValue != ApostropheIdx || int.MaxValue != CommentIdx || int.MaxValue != ScommentIdx || int.MaxValue != SscommentIdx);

            PureContent += FunctionContent;

            FunctionContent = PureContent;

            int bracket = 0;

            int idx = 0;

            Match BracketMatch;

            do
            {
                BracketMatch = Regex.Match(FunctionContent, @"([()])");

                if ("(" == BracketMatch.Value)
                {
                    bracket--;
                }
                else if (")" == BracketMatch.Value)
                {
                    bracket++;
                }

                idx += BracketMatch.Index + BracketMatch.Length;

                FunctionContent = FunctionContent.Substring(BracketMatch.Index + BracketMatch.Length);

            } while (0 != bracket && "" != BracketMatch.Value);

            if (0 == bracket)
            {
                int ProtoID= idx;

                BracketMatch = Regex.Match(FunctionContent, @"([ \t\r\n]*;)");

                if ("" != BracketMatch.Value && 0 == BracketMatch.Index)
                {
                    idx += BracketMatch.Length;
                }
                else
                {
                    bracket = -1;

                    BracketMatch = Regex.Match(FunctionContent, @"([ \t\r\n]*{)");

                    if ("" != BracketMatch.Value && 0 == BracketMatch.Index)
                    {
                        idx += BracketMatch.Length;

                        FunctionContent = FunctionContent.Substring(BracketMatch.Length);

                        do
                        {
                            BracketMatch = Regex.Match(FunctionContent, @"([{}])");

                            if ("{" == BracketMatch.Value)
                            {
                                bracket--;
                            }
                            else if ("}" == BracketMatch.Value)
                            {
                                bracket++;
                            }

                            idx += BracketMatch.Index + BracketMatch.Length;

                            FunctionContent = FunctionContent.Substring(BracketMatch.Index + BracketMatch.Length);

                        } while (0 != bracket && "" != BracketMatch.Value);
                    }
                }

                if (0 == bracket)
                {
                    if (PureContent.Length > idx)
                    {
                        PureContent = PureContent.Remove(idx);
                    }

                    string KeywordPattern = @"\b(_Alignas|_Alignof|_Atomic|_Bool|_Complex|_Generic|_Imaginary|_Noreturn|_Static_assert|_Thread_local|alignas|alignof|and|and_eq|asm|auto|bitand|bitor|bool|break|case|catch|char|char16_t|char32_t|class|compl|const|const_cast|constexpr|continue|decltype|default|delete|do|double|dynamic_cast|else|enum|explicit|export|extern|false|float|for|friend|goto|if|inline|int|long|mutable|namespace|new|noexcept|not|not_eq|nullptr|operator|or|or_eq|private|protected|public|register|reinterpret_cast|restrict|return|short|signed|sizeof|static|static_assert|static_cast|struct|switch|template|this|thread_local|throw|true|try|typedef|typeid|typename|union|unsigned|using|virtual|void|volatile|wchar_t|while|xor|xor_eq)\b";

                    string NoVoidContent = PureContent;

                    while (NoVoidContent != Regex.Replace(NoVoidContent, @"((((\r\n)|(\r)|(\n))|(^))(\s*?)(((\r\n)|(\r)|(\n))|($)))", "\r\n", RegexOptions.Singleline))
                    {
                        NoVoidContent = Regex.Replace(NoVoidContent, @"((((\r\n)|(\r)|(\n))|(^))(\s*?)(((\r\n)|(\r)|(\n))|($)))", "\r\n", RegexOptions.Singleline);
                    }

                    while (NoVoidContent != Regex.Replace(NoVoidContent, @"((^)((\r\n)|(\r)|(\n)))", "", RegexOptions.Singleline))
                    {
                        NoVoidContent = Regex.Replace(NoVoidContent, @"((^)((\r\n)|(\r)|(\n)))", "", RegexOptions.Singleline);
                    }

                    string PrototypeContent = PureContent;

                    PrototypeContent = PrototypeContent.Remove(ProtoID);

                    while (PrototypeContent != Regex.Replace(PrototypeContent, @"((((\r\n)|(\r)|(\n))|(^))(\s*?)(((\r\n)|(\r)|(\n))|($)))", "\r\n", RegexOptions.Singleline))
                    {
                        PrototypeContent = Regex.Replace(PrototypeContent, @"((((\r\n)|(\r)|(\n))|(^))(\s*?)(((\r\n)|(\r)|(\n))|($)))", "\r\n", RegexOptions.Singleline);
                    }

                    while (PrototypeContent != Regex.Replace(PrototypeContent, @"((^)((\r\n)|(\r)|(\n)))", "", RegexOptions.Singleline))
                    {
                        PrototypeContent = Regex.Replace(PrototypeContent, @"((^)((\r\n)|(\r)|(\n)))", "", RegexOptions.Singleline);
                    }

                    PrototypeContent = Regex.Replace(PrototypeContent, @"((\r\n)|(\r)|(\n))", "", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\t)", " ", RegexOptions.Singleline);

                    while (PrototypeContent != Regex.Replace(PrototypeContent, @"(  )", " ", RegexOptions.Singleline))
                    {
                        PrototypeContent = Regex.Replace(PrototypeContent, @"(  )", " ", RegexOptions.Singleline);
                    }

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(,)", ", ", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( ,)", ",", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( \))", ")", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( \()", "(", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\) )", ")", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\( )", "(", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( \})", "}", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( \{)", "{", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\} )", "}", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\{ )", "{", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"( \*)", "*", RegexOptions.Singleline);

                    PrototypeContent = Regex.Replace(PrototypeContent, @"(\* )", "*", RegexOptions.Singleline);

                    StatisticsTable.Add(new StatisticsSet()
                    {
                        FunctionName = PrototypeContent,

                        LinesCount = (Regex.Matches(PureContent, @"((\r\n)|(\r)|(\n))").Count + 1).ToString(),

                        PureLinesCount = (Regex.Matches(NoVoidContent, @"((\r\n)|(\r)|(\n))").Count + 1).ToString(),

                        KeywordCount = Regex.Matches(PureContent, KeywordPattern).Count.ToString(),
                    });
                }
            }
        }

        private void AnalyzeClass(CodeElement ClassCodeElement)
        {
            Dispatcher.VerifyAccess();

            CodeClass ClassElement = ClassCodeElement as CodeClass;

            foreach (CodeElement CurrentCodeElement in ClassElement.Children)
            {
                if (CurrentCodeElement.Kind == vsCMElement.vsCMElementFunction)
                {
                    CodeFunction FunctionElement = CurrentCodeElement as CodeFunction;

                    if (FunctionElement.Access == vsCMAccess.vsCMAccessPublic)
                    {
                        FunctionList.Add(FunctionElement);
                    }
                    else if (FunctionElement.Access == vsCMAccess.vsCMAccessProtected)
                    {
                        FunctionList.Add(FunctionElement);
                    }
                    else if (FunctionElement.Access == vsCMAccess.vsCMAccessPrivate)
                    {
                        FunctionList.Add(FunctionElement);
                    }
                }
            }
        }

        private void AnalyzeNamespace(CodeNamespace NamespaceElement)
        {
            Dispatcher.VerifyAccess();

            foreach (CodeElement CurrentCodeElement in NamespaceElement.Children)
            {
                if (CurrentCodeElement.Kind == vsCMElement.vsCMElementFunction)
                {
                    CodeFunction FunctionElement = CurrentCodeElement as CodeFunction;

                    FunctionList.Add(FunctionElement);
                }
                else if (CurrentCodeElement.Kind == vsCMElement.vsCMElementNamespace)
                {
                    CodeNamespace ChildNamespaceElement = CurrentCodeElement as CodeNamespace;

                    AnalyzeNamespace(ChildNamespaceElement);
                }
            }
        }

        private void AnalyzeCurrentFile(FileCodeModel2 MyFileCodeModel)
        {
            Dispatcher.VerifyAccess();

            foreach (CodeElement CurrentCodeElement in MyFileCodeModel.CodeElements)
            {
                if (CurrentCodeElement.Kind == vsCMElement.vsCMElementNamespace)
                {
                    CodeNamespace CurrentNamespaceElement = CurrentCodeElement as CodeNamespace;

                    AnalyzeNamespace(CurrentNamespaceElement);
                }
                else if (CurrentCodeElement.Kind == vsCMElement.vsCMElementFunction)
                {
                    CodeFunction FunctionElement = CurrentCodeElement as CodeFunction;

                    FunctionList.Add(FunctionElement);
                }
                else if (CurrentCodeElement.Kind == vsCMElement.vsCMElementClass)
                {
                    AnalyzeClass(CurrentCodeElement);
                }
            }
        }

        private void MenuItemCallback()
        {
            Dispatcher.VerifyAccess();

            StatisticsTable.Clear();

            DTE2 MyDte;

            try
            {
                MyDte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
 
                ProjectItem MyProjectItem = MyDte.ActiveDocument.ProjectItem;

                FileCodeModel2 MyFileCodeModel = (FileCodeModel2)MyProjectItem.FileCodeModel;

                if (MyFileCodeModel != null)
                {
                    FunctionList = new List<CodeFunction>();

                    AnalyzeCurrentFile(MyFileCodeModel);

                    for (int i = 0; i < FunctionList.Count; i++)
                    {
                        ParseFunction(i);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MenuItemCallback();

            StatisticsList.Items.Refresh();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MenuItemCallback();

            StatisticsList.Items.Refresh();
        }
    }
}