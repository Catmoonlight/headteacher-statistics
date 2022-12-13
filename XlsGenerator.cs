using Excel = Microsoft.Office.Interop.Excel;

namespace HogStatGenerator
{
    internal class XlsGenerator : IDisposable
    {
        public string fileName;
        private readonly Excel.Application _app;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;
        internal bool isChanged = false;
        internal bool isSumFilled = false;
        internal bool isMarksFilled = false;
        internal bool isRelevenceFilled = false;
        private double[] _marksPercent;
        private int _lastProblemColumn;
        internal int studentsCount;
        internal int activeStudentsCount;
        private int examMax;
        private string _prevYearMarkColumnAddress;
        private string _examMarkColumnAddress;
        private string _examSumColumnAddress;
        internal const string _studentMiss = "отсутсвовал";
        internal const string _prevYearMark = "Отметка за предыдущий год";

        internal XlsGenerator(string fileName, double[] _marksPercent)
        {
            this.fileName = fileName;
            _app = new Excel.Application
            {
                Visible = false,
                SheetsInNewWorkbook = 1
            };
            isChanged = false;
            isSumFilled = false;
            _workbook = _app.Workbooks.Add(fileName);
            _worksheet = (Excel.Worksheet)_workbook.Worksheets.get_Item(1);
            this._marksPercent = _marksPercent;
        }

        private void CalculateConstants()
        {
            
            studentsCount = GetStudentsCount();
            activeStudentsCount = GetActiveStudentsCount();
            _prevYearMarkColumnAddress = GetPrevYearMarkColumnAddress();
            _examMarkColumnAddress = GetExamMarkColumnAddress();
            _examSumColumnAddress = GetScoreColumnAsRangeAddress();
            examMax = GetExamMax();
        }

        internal void Generate()
        {
            //_workbook.Close();
            //throw new NotFiniteNumberException();
            if (isChanged)
                return;
            isChanged = true;
            _lastProblemColumn = GetLastProblemColumn();
            CalculateConstants();
            FillExamSums();
            FillMarks();
            FillRelevence();

            FillWorksheet();
        }

        private void FillWorksheet()
        {
            CalculateStatistic();
            DrawGraphics();
        }

        private void DrawGraphics()
        {
            FillMarksChartRange();
            DrawMarksDiagramm();
            //DrawSumsDiagramm();
        }

        private void CalculateStatistic()
        {
            FillCorrelation();
            FillMedian();
            FillMode();
            FillMean();
        }

        internal bool IsCorrect()
        {
            if (isChanged)
            {
                return true;
            }
            return IsWorkbookFormatCorrect();
        }

        internal void FillExamSums()
        {
            isSumFilled = true;
            _worksheet.Cells[1, _lastProblemColumn + 2].Value = "Итог";
            for (int row = 2; row <= GetStudentsCount() + 1; row++)
            {
                if ((_worksheet.Cells[row, 7].Value) is double)
                {
                    FillExamSumByRow(row);
                }
            }
        }

        internal void FillMarks()
        {
            if (!isSumFilled)
            {
                throw new ArgumentException("Expected filled sum column");
            }
            isMarksFilled = true;
            _worksheet.Cells[1, _lastProblemColumn + 3].Value = "Оценка";
            for (int row = 2; row <= GetStudentsCount() + 1; row++)
            {
                FillMarkByRow(row);
            }
        }

        private void FillRelevence()
        {
            if (!isMarksFilled)
            {
                throw new ArgumentException("Expected filled marks column");
            }
            isRelevenceFilled = true;
            _worksheet.Cells[1, _lastProblemColumn + 4].Value = "соответствие оценок за работу текущим";
            for (int row = 2; row < GetStudentsCount() + 1; row++)
            {
                FillRelevenceByRow(row);
            }
        }

        private void FillCorrelation()
        {
            FillMarksCorrelation();
            FillScoreCorrelation();
        }

        private void FillScoreCorrelation()
        {
            _worksheet.Cells[1, _lastProblemColumn + 5].Value = "Корреляция баллов с годовыми оценками:";
            (_worksheet.Cells[1, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=КОРРЕЛ({_examSumColumnAddress};{_prevYearMarkColumnAddress})";
        }

        private void FillMarksCorrelation()
        {
            _worksheet.Cells[2, _lastProblemColumn + 5].Value = "Корреляция оценки с годовыми оценками:";
            (_worksheet.Cells[2, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=КОРРЕЛ({_examMarkColumnAddress};{_prevYearMarkColumnAddress})";
        }

        private void FillMedian()
        {
            FillMarkMedian();
            FillScoreMedian();
        }

        private void FillScoreMedian()
        {
            _worksheet.Cells[3, _lastProblemColumn + 5].Value = "Медиана баллов:";
            (_worksheet.Cells[3, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=МЕДИАНА({_examSumColumnAddress})";
        }

        private void FillMarkMedian()
        {
            _worksheet.Cells[4, _lastProblemColumn + 5].Value = "Медиана оценок:";
            (_worksheet.Cells[4, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=МЕДИАНА({_examMarkColumnAddress})";
        }

        private void FillMode()
        {
            FillMarkMode();
            FillScoreMode();
        }

        private void FillScoreMode()
        {
            _worksheet.Cells[5, _lastProblemColumn + 5].Value = "Мода баллов:";
            (_worksheet.Cells[5, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=МОДА.НСК({_examSumColumnAddress})";
        }

        private void FillMarkMode()
        {
            _worksheet.Cells[6, _lastProblemColumn + 5].Value = "Мода оценок:";
            (_worksheet.Cells[6, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=МОДА.НСК({_examMarkColumnAddress})";
        }

        private void FillMean()
        {
            FillMarkMean();
            FillScoreMean();
        }

        private void FillScoreMean()
        {
            _worksheet.Cells[7, _lastProblemColumn + 5].Value = "Среднее значение баллов:";
            (_worksheet.Cells[7, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=СРЗНАЧ({_examSumColumnAddress})";
        }

        private void FillMarkMean()
        {
            _worksheet.Cells[8, _lastProblemColumn + 5].Value = "Среднее значение оценок:";
            (_worksheet.Cells[8, _lastProblemColumn + 6] as Excel.Range)!
                .FormulaLocal = $"=СРЗНАЧ({_examMarkColumnAddress})";
        }

        private void DrawMarksDiagramm()
        {
            Excel.Range chartRange = GetMarksChartRange();
            var chartObjects = (Excel.ChartObjects)_worksheet.ChartObjects(Type.Missing);
            Excel.ChartObject chartObject = chartObjects.Add((_lastProblemColumn + 12) * 60, 30, 300, 300);
            Excel.Chart chart = chartObject.Chart;
            chart.SetSourceData(chartRange);
            chart.ChartType = Excel.XlChartType.xlPie;
        }

        private void FillMarksChartRange()
        {
            _worksheet.Cells[1, _lastProblemColumn + 9].Value = "Всего:";
            _worksheet.Cells[1, _lastProblemColumn + 10].Value = activeStudentsCount;
            for (int i = 2; i <= 5; i++)
            {
                _worksheet.Cells[i, _lastProblemColumn + 8].Value = i;
                _worksheet.Cells[i, _lastProblemColumn + 10].FormulaLocal = GetMarkCount(i);
                (_worksheet.Cells[i, _lastProblemColumn + 9] as Excel.Range)!
                    .FormulaLocal =
                    $"={(_worksheet.Cells[i, _lastProblemColumn + 10] as Excel.Range)!.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing)}" +
                    $" / " +
                    $"{(_worksheet.Cells[1, _lastProblemColumn + 10] as Excel.Range)!.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing)}";
                (_worksheet.Cells[i, _lastProblemColumn + 9] as Excel.Range)!.NumberFormat = "0.00%;[Red]-0.00%";
            }
        }

        private Excel.Range GetMarksChartRange()
        {
            return _worksheet.get_Range((_worksheet.Cells[2, _lastProblemColumn + 8] as Excel.Range)!, (_worksheet.Cells[5, _lastProblemColumn + 9] as Excel.Range)!);
        }

        private string GetMarkCount(int i)
        {
            return $"=СЧЕТЕСЛИ({_examMarkColumnAddress};{i})";
        }

        private void DrawSumsDiagramm()
        {
            throw new NotImplementedException();
        }

        private void FillRelevenceByRow(int row)
        {
            string lastYearMarkCellAsRangeAdress = (_worksheet.Cells[row, _lastProblemColumn + 1] as Excel.Range)!
                                                   .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            string examMarkCellAsRangeAdress = (_worksheet.Cells[row, _lastProblemColumn + 3] as Excel.Range)!
                                                   .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            string greater = $"=ЕСЛИ({examMarkCellAsRangeAdress} > {lastYearMarkCellAsRangeAdress}; \"Выше\";";
            string less = $"ЕСЛИ({examMarkCellAsRangeAdress} < {lastYearMarkCellAsRangeAdress}; \"Ниже\"; \"Соотв\"";
            Excel.Range relevence = (_worksheet.Cells[row, _lastProblemColumn + 4] as Excel.Range)!;
            relevence.FormulaLocal = greater + less + "))";
        }

        private void FillMarkByRow(int row)
        {
            string sumCellAsRangeAdress = (_worksheet.Cells[row, _lastProblemColumn + 2] as Excel.Range)!
                                          .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            Excel.Range markCellAsRange = (_worksheet.Cells[row, _lastProblemColumn + 3] as Excel.Range)!;
            var examPercent = $"{sumCellAsRangeAdress} / {examMax}";
            string FMark = $"=ЕСЛИ({examPercent} < {_marksPercent[0]}; 2;";
            string CMark = $"ЕСЛИ({examPercent} < {_marksPercent[1]}; 3;";
            string BMark = $"ЕСЛИ({examPercent} < {_marksPercent[2]}; 4; 5";
            markCellAsRange.FormulaLocal = FMark + CMark + BMark + ")))";

        }

        private bool IsWorkbookFormatCorrect() => IsStudentsInfoCorrect() && IsProblemsInfoCorrect();

        private bool IsProblemsInfoCorrect()
        {
            int column = 7;
            while (((string)(_worksheet.Cells[1, column])).Contains("б)"))
            {
                var scoreInParentheses = ((string)(_worksheet.Cells[0, column])).Split().Last();
                if (!(scoreInParentheses.Contains('(') && scoreInParentheses.Contains("б)")))
                {
                    return false;
                }
                if (!scoreInParentheses.Replace("(", "").Replace("б)", "").IsNumber())
                {
                    return false;
                }
                column++;
            }
            return !(column == 7)
                && (((string)(_worksheet.Cells[0, column])) == _prevYearMark);
        }

        private bool IsStudentsInfoCorrect()
        {
            return ((string)(_worksheet.Cells[1, 1])).ToLowerRus() == "ф.и."
                   && ((string)(_worksheet.Cells[1, 2])).ToLowerRus() == "пол"
                   && ((string)(_worksheet.Cells[1, 3])).ToLowerRus() == "класс"
                   && ((string)(_worksheet.Cells[1, 4])).ToLowerRus() == "наименование класса"
                   && ((string)(_worksheet.Cells[1, 5])).ToLowerRus() == "код"
                   && ((string)(_worksheet.Cells[1, 6])).ToLowerRus() == "вариант";
        }

        private int GetStudentsCount()
        {
            var row = 2;
            var cnt = 0;
            while ((_worksheet.Cells[row, 5].Value) is double)
            {
                row++;
                cnt++;
            }
            return cnt;
        }

        private int GetActiveStudentsCount()
        {
            var row = 2;
            var cnt = 0;
            while ((_worksheet.Cells[row, 5].Value) is double)
            {
                cnt++;
                if (!((_worksheet.Cells[row, 7].Value) is double))
                {
                    cnt--;
                }
                row++;
            }
            return cnt;
        }

        private void FillExamSumByRow(int row)
        {
            int sumColumn = _lastProblemColumn + 2;
            Excel.Range cellAsRange = (_worksheet.Cells[row, sumColumn] as Excel.Range)!;
            cellAsRange!.FormulaLocal = $"=СУММ({GetProblemRangeByRow(row)})";
        }

        private string GetProblemRangeByRow(int row)
        {
            Excel.Range problemRange = (Excel.Range)((_worksheet as Excel.Worksheet).Range[_worksheet.Cells[row, 7], _worksheet.Cells[row, _lastProblemColumn]]);
            return (string)problemRange.Address[1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing];
        }

        private int GetLastProblemColumn()
        {
            int column = 7;
            while ((string)(_worksheet.Cells[1, column].Value) != _prevYearMark)
            {
                column++;
            }
            return column - 1;
        }

        private string GetScoreColumnAsRangeAddress()
        {
            if (!isChanged)
            {
                throw new ArgumentException("Score sum have to be filled");
            }
            return _worksheet.get_Range(_worksheet.Cells[1, _lastProblemColumn + 2] as Excel.Range
                                       , _worksheet.Cells[studentsCount, _lastProblemColumn + 2] as Excel.Range)
                             .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }

        private string GetPrevYearMarkColumnAddress()
        {
            return _worksheet.get_Range(_worksheet.Cells[1, _lastProblemColumn + 1] as Excel.Range
                                       , _worksheet.Cells[studentsCount, _lastProblemColumn + 1] as Excel.Range)
                             .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }

        private string GetExamMarkColumnAddress()
        {
            if (!isChanged)
            {
                throw new ArgumentException("Marks have to be filled");
            }
            return _worksheet.get_Range(_worksheet.Cells[1, _lastProblemColumn + 3] as Excel.Range
                                       , _worksheet.Cells[studentsCount, _lastProblemColumn + 3] as Excel.Range)
                             .get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }

        private int GetExamMax()
        {
            int sum = 0;
            var column = 7;
            while ((string)(_worksheet.Cells[1, column].Value) != _prevYearMark)
            {
                sum += GetMaxProblemScore((string)(_worksheet.Cells[1, column].Value));
                column++;
            }
            return sum;
        }

        private int GetMaxProblemScore(string problemName)
        {
            Console.WriteLine(problemName.Split()[1].Replace("(", "").Replace("б)", ""));
            return Int32.Parse(problemName.Split()[1].Replace("(", "").Replace("б)", ""));
        }

        public void Dispose()
        {
            _workbook.SaveAs2(fileName + "_");
            _workbook.Close();
            _app.Quit();
        }
    }
}
