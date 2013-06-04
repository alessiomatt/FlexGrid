using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using C1.Win.C1FlexGrid;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace Sample
{
    public static class FlexGridExtension
    {
        #region 초기화

        // 스타일 초기화
        public static void x_InitializeStyle(this C1FlexGrid flexGrid)
        {
            flexGrid.Styles.Clear();

            //CellStyle추가
            CellStyle cs = flexGrid.Styles.Add(StyleType.DefaultCell.ToString());
            cs.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0);
            cs.BackColor = Color.White;

            //Row 추가
            cs = flexGrid.Styles.Add(StyleType.NewRow.ToString());
            cs.ForeColor = System.Drawing.ColorTranslator.FromHtml("#015fd4");
            cs.Font = new System.Drawing.Font(flexGrid.Font, FontStyle.Bold);

            //Row 변경
            cs = flexGrid.Styles.Add(StyleType.ChangeRow.ToString());
            cs.ForeColor = System.Drawing.ColorTranslator.FromHtml("#b8154a");
            cs.Font = new System.Drawing.Font(flexGrid.Font, FontStyle.Bold);

            //GrandTotalCellStyle
            cs = flexGrid.Styles.Add("GrandTotalCellStyle");
            cs.TextAlign = TextAlignEnum.RightCenter;

            //삭제 데이터
            cs = flexGrid.Styles.Add(StyleType.DeleteRow.ToString());
            cs.Font = new System.Drawing.Font(flexGrid.Font, FontStyle.Strikeout);

            //CellStyle추가
            cs = flexGrid.Styles.Add(StyleType.ReadOnlyCell.ToString());
            cs.BackColor = System.Drawing.ColorTranslator.FromHtml("#eeeeee");

            // 빈영역
            flexGrid.Styles.EmptyArea.BackColor = System.Drawing.ColorTranslator.FromHtml("#fff8f8f8");

            // Color 관련
            flexGrid.Styles.Normal.ForeColor = System.Drawing.ColorTranslator.FromHtml("#ff3d3d3d");
            flexGrid.Styles.Normal.Border.Color = System.Drawing.ColorTranslator.FromHtml("#ffcccccc");

            // 그리드 헤더 배경색
            flexGrid.Styles.Fixed.BackColor = System.Drawing.ColorTranslator.FromHtml("#f4f4f4");
            flexGrid.Styles.Fixed.ForeColor = System.Drawing.ColorTranslator.FromHtml("#4A4A4A");
            flexGrid.Styles.Fixed.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            flexGrid.Styles.Fixed.Border.Color = Color.Black;

            // 틀고정 색상
            flexGrid.Styles.Frozen.BackColor = System.Drawing.Color.White;

            // HighLight
            flexGrid.Styles.Highlight.BackColor = Color.FromArgb(255, 229, 155);
            flexGrid.Styles.Highlight.ForeColor = System.Drawing.ColorTranslator.FromHtml("#000000");

            flexGrid.Styles.Focus.BackColor = flexGrid.Styles.Highlight.BackColor;
            flexGrid.Styles.Focus.ForeColor = flexGrid.Styles.Highlight.ForeColor;

            // Grand Total
            flexGrid.Styles[CellStyleEnum.GrandTotal].BackColor = Color.FromArgb(84, 91, 136);
            flexGrid.Styles[CellStyleEnum.GrandTotal].ForeColor = Color.FromArgb(224, 225, 235);
            flexGrid.Styles[CellStyleEnum.GrandTotal].TextAlign = TextAlignEnum.RightCenter;

            flexGrid.Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter;


            flexGrid.OwnerDrawCell += delegate(object sender, OwnerDrawCellEventArgs e)
            {
                if (e.Row < flexGrid.Rows.Fixed || e.Col < flexGrid.Cols.Fixed)
                    return;

                int index = flexGrid.Rows[e.Row].DataIndex;
                if (index < 0 || flexGrid.DataSource == null) return;
                CurrencyManager cm = (CurrencyManager)flexGrid.BindingContext[flexGrid.DataSource, flexGrid.DataMember];

                if (cm.List.Count < index)
                    return;

                DataRowView drv = cm.List[index] as DataRowView;

                // select style based on row state
                switch (drv.Row.RowState)
                {
                    case DataRowState.Added:
                        e.Style = flexGrid.Styles[StyleType.NewRow.ToString()];
                        break;
                    case DataRowState.Modified:
                        e.Style = flexGrid.Styles[StyleType.ChangeRow.ToString()];
                        break;
                    case DataRowState.Deleted:
                        e.Style = flexGrid.Styles[StyleType.DeleteRow.ToString()];
                        break;
                    case DataRowState.Unchanged:
                        break;
                    default:
                        break;
                }
            };
        }


        // 디자인 초기화
        /// <summary>
        /// 그리드를 표준에 맞게 기본설정을 초기화
        /// </summary>
        public static void x_InitializeDesign(this C1FlexGrid flexGrid)
        {
            // 기본 컬럼의 폭 값이 0 이면 조정해 준다.
            if (flexGrid.Cols.DefaultSize == 0)
                flexGrid.Cols.DefaultSize = 85;

            // Default로 모두 읽기 전용으로 만든다.
            // this.AllowEditing = false;  를 사용하면 부분적 컬럼수정 기능이 않됨.
            flexGrid.AutoResize = false;
            flexGrid.AutoGenerateColumns = false;
            flexGrid.DrawMode = DrawModeEnum.OwnerDraw;
            flexGrid.AutoClipboard = true;
            flexGrid.AllowFreezing = AllowFreezingEnum.Both;

            // Border 스타일
            flexGrid.BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.FixedSingle;

            // Define Default Font
            flexGrid.Font = new System.Drawing.Font("돋움", 9);

            // Selection Mode
            flexGrid.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.RowRange;

            flexGrid.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveDown;
            flexGrid.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcrossOut;
            flexGrid.Cols.Count = 1;
            flexGrid.Rows.Count = 1;
        }

        #endregion

        #region 데이터 바인드
        /// <summary>
        /// 데이터 바인드(DataSource 속성을 사용하지 않는다.)
        /// </summary>
        /// <param name="view"></param>
        public static void x_DataBind(this C1FlexGrid flexGrid, System.Data.DataView view, bool viewDeleteRow)
        {
            if (viewDeleteRow)
            {
                view.RowStateFilter = view.RowStateFilter | DataViewRowState.Deleted;
            }

            flexGrid.DataSource = view;
            //ClearCellEditable();

            view.ListChanged += delegate(object sender, System.ComponentModel.ListChangedEventArgs e)
            {
                if (e.NewIndex == -1)
                {
                    SetNumbering(flexGrid.Rows.Fixed, flexGrid);
                }
                else
                {
                    SetNumbering(e.NewIndex + flexGrid.Rows.Fixed, flexGrid);
                }
            };

            view.Table.RowChanged += delegate(object sender, DataRowChangeEventArgs e)
            {
                if (e.Action == DataRowAction.Add)
                    flexGrid.Invalidate();
            };

            view.Table.ColumnChanged += delegate(object sender, DataColumnChangeEventArgs e)
            {
                if (e.Row.RowState != DataRowState.Detached &&
                    (e.Row[e.Column.ColumnName, DataRowVersion.Current] == DBNull.Value
                    || e.Row[e.Column.ColumnName, DataRowVersion.Current] == null))
                {
                    if (e.ProposedValue == null
                        || Convert.ToString(e.ProposedValue).Trim() == string.Empty)
                    {
                        e.Row.CancelEdit();
                        return;
                    }
                }

                e.Row.EndEdit();
            };

            SetNumbering(flexGrid.Rows.Fixed, flexGrid);
        }

        // Row Number를 붙인다.
        private static void SetNumbering(int rowIndex, C1FlexGrid flexGrid)
        {
            if (flexGrid.DataSource != null)
            {
                for (int i = rowIndex; i < flexGrid.Rows.Count; i++)
                {
                    flexGrid[i, 0] = i - flexGrid.Rows.Fixed + 1;
                }
            }
        }

        #endregion

        #region 컬럼 추가

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, bool hidden)
        {
            x_AddColumn(flexGrid, colName, colName, hidden);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption)
        {
            x_AddColumn(flexGrid, colName, caption, false);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, 100, TextAlignEnum.LeftCenter);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, TextAlignEnum colAlign)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, colAlign, false);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, Type columnDataType, TextAlignEnum colAlign)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, columnDataType, colAlign, false);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, TextAlignEnum colAlign, bool readOnly)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, colAlign, readOnly, "");
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, Type columnDataType, TextAlignEnum colAlign, bool readOnly)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, columnDataType, colAlign, readOnly, "");
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, TextAlignEnum colAlign, bool readOnly, string format)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, TextAlignEnum.CenterCenter, colAlign, readOnly, format);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, Type columnDataType, TextAlignEnum colAlign, bool readOnly, string format)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, columnDataType, TextAlignEnum.CenterCenter, colAlign, readOnly, format);
        }

        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden, int width, TextAlignEnum headerAlign,
            TextAlignEnum colAlign, bool readOnly, string format)
        {
            x_AddColumn(flexGrid, colName, caption, hidden, width, null, headerAlign, colAlign, readOnly, format);
        }


        public static void x_AddColumn(this C1FlexGrid flexGrid, string colName, string caption, bool hidden
            , int width, Type columnDataType, TextAlignEnum headerAlign, TextAlignEnum colAlign
            , bool readOnly, string format)
        {
            Column column = flexGrid.Cols.Add();
            column.Name = colName;
            column.Caption = caption;

            if (readOnly)
            {
                column.AllowEditing = false;
                //column.Style = flexGrid.GetStyle(StyleType.ReadOnlyCell);
            }
            else
            {
                column.AllowEditing = true;
                //column.Style = flexGrid.GetStyle(StyleType.DefaultCell);
            }

            column.TextAlign = colAlign;
            column.TextAlignFixed = headerAlign;

            if (columnDataType != null)
            {
                column.DataType = columnDataType;
            }

            column.Visible = !hidden;
            column.Width = width;

            if (!string.IsNullOrEmpty(format))
            {
                column.Format = format;
            }
        }
        #endregion

        #region 컨트롤 컬럼 추가

        /// <summary>
        /// 컬럼 추가 Edit Control 지정
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="caption"></param>
        /// <param name="control"></param>
        public static void x_AddColumnControl(this C1FlexGrid flexGrid, string colName, string caption, Control control)
        {
            x_AddColumnControl(flexGrid, colName, caption, -1, null, control);
        }
        /// <summary>
        /// 컬럼 추가 Edit Control 지정
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="caption"></param>
        /// <param name="width"></param>
        /// <param name="control"></param>
        public static void x_AddColumnControl(this C1FlexGrid flexGrid, string colName, string caption, int width, Control control)
        {
            x_AddColumnControl(flexGrid, colName, caption, width, null, control);
        }

        /// <summary>
        /// 컬럼 추가 Edit Control 지정
        /// </summary>
        /// <param name="colName">컬럼명</param>
        /// <param name="caption">컬럼해더</param>
        /// <param name="width">컬럼 길이</param>
        /// <param name="control">컬럼 바인딩 컨트롤</param>
        /// <param name="format">컬럼 보기 형식</param>
        public static void x_AddColumnControl(this C1FlexGrid flexGrid, string colName, string caption, int width, string format, Control control)
        {
            x_AddColumnControl(flexGrid, colName, caption, width, format, TextAlignEnum.LeftCenter, false, control);
        }

        /// <summary>
        /// 컬럼 추가 Edit Control 지정
        /// </summary>
        /// <param name="colName">컬럼 명</param>
        /// <param name="caption">해더명</param>
        /// <param name="width">길이</param>
        /// <param name="colAlign">데이터 정렬</param>
        /// <param name="readOnly">읽기 전용</param>
        /// <param name="control">컬럼 바인딩 컨트롤</param>
        /// <param name="format">컬럼 보기 형식</param>
        public static void x_AddColumnControl(this C1FlexGrid flexGrid, string colName, string caption, int width, string format
            , TextAlignEnum colAlign, bool readOnly, Control control)
        {
            Column column = flexGrid.Cols.Add();

            column.Caption = caption;
            column.Name = colName;

            if (width != -1)
            {
                column.Width = width;
            }

            if (readOnly)
            {
                column.AllowEditing = false;
                column.Style = flexGrid.Styles[StyleType.ReadOnlyCell.ToString()];
            }
            else
            {
                column.AllowEditing = true;
                column.Style = flexGrid.Styles[StyleType.DefaultCell.ToString()];
            }

            // 컬럼 스타일 명 변경
            string styleName = Guid.NewGuid().ToString("N");
            CellStyle style = flexGrid.Styles.Add(styleName, column.Style);
            column.Style = style;

            column.TextAlign = colAlign;
            column.AllowEditing = !readOnly;

            if (!string.IsNullOrEmpty(format))
            {
                column.Format = format;
            }

            column.Editor = control;
        }

        #endregion

        #region 기타 조작

        /// <summary>
        /// 열 추가
        /// </summary>
        /// <param name="flexGrid"></param>
        /// <param name="rowItems"></param>
        /// <returns></returns>
        public static Row x_AddRow(this C1FlexGrid flexGrid, params object[] rowItems)
        {
            return flexGrid.x_AddRow(false, rowItems);
        }
        /// <summary>
        /// 열 추가
        /// </summary>
        /// <param name="rowItems"></param>
        /// <returns></returns>
        public static Row x_AddRow(this C1FlexGrid flexGrid, bool allowNewRowEditing, params object[] rowItems)
        {
            DataView view = flexGrid.DataSource as DataView;
            if (view != null)
            {
                if (rowItems == null)
                {
                    rowItems = new object[] { };
                }
                DataRow row = view.Table.Rows.Add(rowItems);
                int index = view.Table.Rows.IndexOf(row);
                index += flexGrid.Rows.Fixed;

                flexGrid.Select(index, flexGrid.Col, true);

                if (allowNewRowEditing)
                {
                    flexGrid.x_SetRowStyle(StyleType.DefaultCell, index);
                }

                return flexGrid.Rows[index];
            }

            return null;
        }
        public static Row x_InsertRow(this C1FlexGrid flexGrid, params object[] rowItems)
        {
            return flexGrid.x_InsertRow(false, rowItems);
        }

        /// <summary>
        /// 열 삽입
        /// </summary>
        /// <param name="rowItems"></param>
        /// <returns></returns>
        public static Row x_InsertRow(this C1FlexGrid flexGrid, bool allowNewRowEditing, params object[] rowItems)
        {
            DataView view = flexGrid.DataSource as DataView;
            if (flexGrid.Rows.Count == flexGrid.Rows.Fixed)
            {
                flexGrid.x_AddRow(true, rowItems);
            }
            else if (view != null && flexGrid.Row > -1)
            {
                if (rowItems == null)
                {
                    rowItems = new object[] { };
                }

                DataRow dataRow = view.Table.NewRow();
                dataRow.ItemArray = rowItems;
                view.Table.Rows.InsertAt(dataRow, flexGrid.Row);

                int index = view.Table.Rows.IndexOf(dataRow);
                index += flexGrid.Rows.Fixed;

                flexGrid.Select(index, flexGrid.Col, true);

                if (allowNewRowEditing)
                {
                    flexGrid.x_SetRowStyle(StyleType.DefaultCell, index);
                }

                return flexGrid.Rows[index];
            }

            return null;
        }

        /// <summary>
        /// 선택된 Data 삭제
        /// </summary>
        public static void x_DeleteRow(this C1FlexGrid flexGrid)
        {
            int r1 = flexGrid.Selection.r1;
            int r2 = flexGrid.Selection.r2;

            for (int i = r2; i >= r1; i--)
            {
                DataRowView view = flexGrid.Rows[i].DataSource as DataRowView;
                view.Delete();
            }

            flexGrid.Refresh();
        }

        /// <summary>
        /// 특정 Cell 의 Style 변경
        /// </summary>
        /// <param name="styleType"></param>
        /// <param name="cellRange"></param>
        public static void x_SetCellStyle(this C1FlexGrid flexGrid, StyleType styleType, CellRange cellRange)
        {
            CellStyle style = flexGrid.x_GetStyle(styleType);
            cellRange.Style = style;
        }

        /// <summary>
        /// 특정 컬럼의 스타일 변경
        /// </summary>
        /// <param name="styleType"></param>
        /// <param name="columnName"></param>
        public static void x_SetColumnStyle(this C1FlexGrid flexGrid, StyleType styleType, string columnName)
        {
            int columnIndex = flexGrid.Cols[columnName].Index;

            CellRange cellRange = flexGrid.GetCellRange(flexGrid.Rows.Fixed, columnIndex, flexGrid.Rows.Count - 1, columnIndex);
            flexGrid.x_SetCellStyle(styleType, cellRange);
        }

        /// <summary>
        /// Grid 열의 스타일을 지정한다.
        /// </summary>
        /// <param name="styleType"></param>
        /// <param name="rowIndex"></param>
        public static void x_SetRowStyle(this C1FlexGrid flexGrid, StyleType styleType, int rowIndex)
        {
            CellRange cellRange = flexGrid.GetCellRange(rowIndex, flexGrid.Cols.Fixed, rowIndex, flexGrid.Cols.Count - 1);
            flexGrid.x_SetCellStyle(styleType, cellRange);
        }

        /// <summary>
        /// Grid 내부에 등록 된 스타일 반환
        /// </summary>
        /// <param name="styleType"></param>
        /// <param name="styleName"></param>
        /// <returns></returns>
        public static CellStyle x_GetStyle(this C1FlexGrid flexGrid, StyleType styleType, string styleName)
        {
            return flexGrid.Styles.Add(styleName, styleType.ToString());
        }

        /// <summary>
        /// Grid 내부에 등록 된 스타일 반환
        /// </summary>
        /// <param name="styleType"></param>
        /// <returns></returns>
        public static CellStyle x_GetStyle(this C1FlexGrid flexGrid, StyleType styleType)
        {
            string styleName = Guid.NewGuid().ToString("N");
            return flexGrid.x_GetStyle(styleType, styleName);
        }

        /// <summary>
        /// 특정 컬럼의 데이터에 UnderLine 표시
        /// </summary>
        /// <param name="columnName"></param>
        public static void x_SetLinkedColumn(this C1FlexGrid flexGrid, string columnName)
        {
            flexGrid.Cols[columnName].StyleFixed.Font =
                new System.Drawing.Font(flexGrid.Cols[columnName].StyleFixed.Font, FontStyle.Underline);
        }

        /// <summary>
        /// 선택 영역 Cell Data를 Clipboard 에 복사
        /// </summary>
        public static void x_Copy(this C1FlexGrid flexGrid)
        {
            Clipboard.SetText(flexGrid.Selection.Clip);
        }

        /// <summary>
        /// 선택된 Cell 에 Clipboard의 데이터를 복사
        /// </summary>
        public static void x_Paste(this C1FlexGrid flexGrid)
        {
            string pasteText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(pasteText))
            {
                int row = -1;
                int col = -1;
                pasteText.Replace("\n", "");
                string[] rows = pasteText.Split('\r');
                row = flexGrid.Row + rows.Length - 1;
                if (rows.Length > 0)
                {
                    string[] columns = rows[0].Split('\t');
                    col = flexGrid.Col + columns.Length - 1;

                    if (row > flexGrid.Rows.Count)
                    {
                        row = flexGrid.Rows.Count;
                    }

                    if (col > flexGrid.Cols.Count)
                    {
                        col = flexGrid.Cols.Count;
                    }

                    CellRange cellRange = flexGrid.GetCellRange(flexGrid.Row, flexGrid.Col, row, col);
                    cellRange.Clip = pasteText.Replace("\\n", "");
                }
            }
        }

        /// <summary>
        /// 그리드 입력 가능 키값 체크
        /// </summary>
        /// <param name="keyCode"></param>
        /// <returns></returns>
        public static bool x_CheckInputKey(this C1FlexGrid flexGrid, Keys keyCode)
        {
            Keys charKey = keyCode ^ Keys.Shift;

            if (keyCode >= Keys.A && keyCode <= Keys.Z)
            {
                return true;
            }
            else if (charKey >= Keys.A && charKey <= Keys.Z)
            {
                return true;
            }
            else if (keyCode >= Keys.D0 && keyCode <= Keys.D9)
            {
                return true;
            }
            else if (keyCode >= Keys.Oem1 & keyCode <= Keys.OemBackslash)
            {
                return true;
            }
            else if (keyCode == Keys.Enter
                || keyCode == Keys.F2)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Flex Grid 를 Excel 파일로 저장
        /// </summary>
        /// <param name="excelFilePathorName">파일 전체 경로 또는 파일명</param>
        /// <param name="sheetName">Excel Sheet명</param>
        /// <param name="showExcel">저장 후 Excel 실행 유무</param>
        /// <returns>저장 경로 반환</returns>
        public static string x_SaveToExcel(this C1FlexGrid flexGrid, string excelFilePathorName, string sheetName, bool showExcel)
        {
            return flexGrid.x_SaveToExcel(excelFilePathorName, sheetName, showExcel, false);
        }

        /// <summary>
        /// Flex Grid 를 Excel 파일로 저장
        /// </summary>
        /// <param name="excelFilePathorName">파일 전체 경로 또는 파일명</param>
        /// <param name="sheetName">Excel Sheet명</param>
        /// <param name="showExcel">저장 후 Excel 실행 유무</param>
        /// <param name="showHideColumn">Grid 숨김 컬럼 표시 여부</param>
        /// <returns>저장 경로 반환</returns>
        public static string x_SaveToExcel(this C1FlexGrid flexGrid, string excelFilePathorName, string sheetName, bool showExcel, bool showHideColumn)
        {
            string excelFilePath = excelFilePathorName;

            if (string.IsNullOrEmpty(excelFilePath))
            {
                // 임시 파일 경로 생성
                excelFilePath = Path.GetTempFileName();
                File.Delete(excelFilePath);
            }

            excelFilePath = RemoveExtension(excelFilePath);
            excelFilePath = string.Format("{0}_{1:yyyyMMdd}", excelFilePath, DateTime.Now);

            if (!Path.IsPathRooted(excelFilePath))
            {
                // 상대 경로인 경우
                string tempPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\DPMS";
                if (!Directory.Exists(tempPath))
                {
                    Directory.CreateDirectory(tempPath);
                }
                string tempFilePath = Path.Combine(tempPath, excelFilePath) + ".xls";

                for (int index = 0; ; index++)
                {
                    if (!File.Exists(tempFilePath))
                    {
                        // 파일이 존재 하지 않는 경우
                        excelFilePath = tempFilePath;
                        break;
                    }

                    string tempFileName = string.Format("{0}_{1}.xls", excelFilePath, index);
                    tempFilePath = Path.Combine(tempPath, tempFileName);
                }
            }

            if (showHideColumn)
            {
                // 숨김 Column 까지 보임
                flexGrid.SaveExcel(excelFilePath, sheetName, C1.Win.C1FlexGrid.FileFlags.IncludeFixedCells | FileFlags.IncludeMergedRanges | FileFlags.AsDisplayed | FileFlags.SaveMergedRanges);
            }
            else
            {
                flexGrid.SaveExcel(excelFilePath, sheetName, C1.Win.C1FlexGrid.FileFlags.VisibleOnly | FileFlags.IncludeFixedCells | FileFlags.IncludeMergedRanges | FileFlags.SaveMergedRanges | FileFlags.AsDisplayed);
            }

            if (showExcel)
            {
                System.Diagnostics.Process.Start("excel.exe", "\"" + excelFilePath + "\"");
            }

            return excelFilePath;
        }

        /// <summary>
        /// 컬럼에 대한 자동 병합 기능의 적용/비적용 설정
        /// </summary>
        /// <param name="flexGrid">C1FlexGrid 컨트롤</param>
        /// <param name="colsIndex">자동 병합 기능의 적용/비적용할 컬럼의 인덱스 배열</param>
        public static void x_AllowMerging(this C1FlexGrid flexGrid, int[] colsIndex)
        {
            x_AllowMerging(flexGrid, colsIndex, true);
        }

        /// <summary>
        /// 컬럼에 대한 자동 병합 기능의 적용/비적용 설정
        /// </summary>
        /// <param name="flexGrid">C1FlexGrid 컨트롤</param>
        /// <param name="colsIndex">자동 병합 기능의 적용/비적용할 컬럼의 인덱스 배열</param>
        /// <param name="bMerge">자동 병합 여부</param>
        public static void x_AllowMerging(this C1FlexGrid flexGrid, int[] colsIndex, bool bMerge)
        {
            x_AllowMerging(flexGrid, colsIndex, null, bMerge);
        }

        /// <summary>
        /// 컬럼에 대한 자동 병합 기능의 적용/비적용 설정
        /// </summary>
        /// <param name="flexGrid">C1FlexGrid 컨트롤</param>
        /// <param name="colsIndex">자동 병합 기능의 적용/비적용할 컬럼의 인덱스 배열</param>
        /// <param name="rowsIndex">자동 병합 기능의 적용/비적용할 행의 인덱스 배열</param>
        /// <param name="bMerge">자동 병합 여부</param>
        public static void x_AllowMerging(this C1FlexGrid flexGrid, int[] colsIndex, int[] rowsIndex, bool bMerge)
        {
            if (colsIndex != null)
            {
                foreach (int colIndex in colsIndex)
                {
                    x_AllowMerging(flexGrid, colIndex, true, bMerge);
                }
            }

            if (rowsIndex != null)
            {
                foreach (int rowIndex in rowsIndex)
                {
                    x_AllowMerging(flexGrid, rowIndex, false, bMerge);
                }
            }
        }

        /// <summary>
        /// 컬럼에 대한 자동 병합 기능의 적용/비적용 설정
        /// </summary>
        /// <param name="flexGrid">C1FlexGrid 컨트롤</param>
        /// <param name="iIndex">자동 병합 기능의 적용/비적용할 인덱스</param>
        /// <param name="bColumn">컬럼인지 행인지의 여부, 컬럼이면 true이다.</param>
        /// <param name="bMerge">자동 병합 여부</param>
        public static void x_AllowMerging(this C1FlexGrid flexGrid, int iIndex, bool bColumn, bool bMerge)
        {
            // 컬럼인 경우
            if (bColumn)
            {
                if (flexGrid.Cols[iIndex] != null)
                {
                    flexGrid.Cols[iIndex].AllowMerging = bMerge;
                }
            }
            else
            {
                if (flexGrid.Rows[iIndex] != null)
                {
                    flexGrid.Rows[iIndex].AllowMerging = bMerge;
                }
            }
        }

        #endregion

        #region 트리 관련

        /// <summary>
        /// FlexGrid를 트리 형식으로 구성한다.
        /// e.g.) felxGrid.x_TreeSetting(datasource, 0, 1, false, false, false);
        /// </summary>
        /// <param name="flexGrid"></param>
        /// <param name="dt">바인딩을 위한 데이터 소스(데이터 소스는 표현될 트리와 동일한 형식으로 구성되어 있어야 함. 트리 노드에 나타날 컬럼과 현재 노드의 레벨이 반드시 지정되어 있어야 함</param>
        /// <param name="treeCol">트리노드에 나타날 컬럼의 인덱스</param>
        /// <param name="levelCol">노드의 레벨 정보를 가지고 있는 컬럼의 인덱스</param>
        /// <param name="expandLevel">처음 로드시 확장될 레벨 값 (e.g. 0으로 지정된 경우 최상위 노드만 보여지고 나머지 레벨은 접혀있음)</param>
        /// <param name="isVertical">트리를 수직으로 표현할 지를 결정. 즉 노드의 수평 라인을 제거</param>
        /// <param name="isHeader">헤더 표시 여부</param>
        /// <param name="isTopLevel">최상위 노드 표시 여부</param>
        /// <remarks>
        /// -------------------------------------
        /// | treeCol     | level | etc1 | etc2 |
        /// -------------------------------------
        /// | 대분류 Data |   0   | etc1 | etc2 |
        /// -------------------------------------
        /// | 소분류 Data |   1   | etc1 | etc2 |
        /// -------------------------------------
        /// | 중분류 Data |   2   | etc1 | etc2 |
        /// -------------------------------------
        /// 
        /// e.g.) flexGrid.x_TreeSetting(datasource, 0, 1, false, false, false);
        /// </remarks>
        public static void x_TreeSetting(this C1FlexGrid flexGrid, System.Data.DataTable dt, int treeCol, int levelCol, int expandLevel, bool isVertical, bool isHeader, bool isTopLevel)
        {
            int START_ROW = 0;
            int END_ROW = 0;

            treeCol += 10;
            levelCol += 10;
            if (dt.Rows.Count == 0) return;

            flexGrid.BeginUpdate();

            //Tree 초기화(ROW선 없앰)
            if (isVertical)
            {
                C1.Win.C1FlexGrid.CellStyle cs = flexGrid.Styles.Normal;
                cs.Border.Direction = C1.Win.C1FlexGrid.BorderDirEnum.Vertical;
            }

            flexGrid.Tree.Column = treeCol;
            flexGrid.Tree.Style = TreeStyleFlags.Simple;
            flexGrid.AllowMerging = AllowMergingEnum.Nodes;
            flexGrid.Tree.Indent = 1;

            //최상위레벨 존재하는 경우
            if (isTopLevel)
            {
                flexGrid.Rows.Count = dt.Rows.Count + flexGrid.Rows.Fixed + 1;
                START_ROW = 1;
                END_ROW = dt.Rows.Count + 1;
                flexGrid.SetData(flexGrid.Rows.Fixed, levelCol, 0);
                flexGrid.SetData(flexGrid.Rows.Fixed, treeCol, "[전체]");
                flexGrid.Rows[flexGrid.Rows.Fixed].IsNode = true;
                flexGrid.Rows[flexGrid.Rows.Fixed].Node.Level = 0;
            }
            else
            {
                flexGrid.Rows.Count = dt.Rows.Count + flexGrid.Rows.Fixed;
                START_ROW = 0;
                END_ROW = dt.Rows.Count;
            }

            for (int i = START_ROW; i < END_ROW; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (flexGrid.Cols.Count == 0)
                    {
                        flexGrid.SetData(i + flexGrid.Rows.Fixed, j, dt.Rows[i - START_ROW][j]);
                    }
                    else if (flexGrid.Cols.Count > 0 && FlexGridFindCol(flexGrid, dt.Columns[j].ColumnName) > -1)
                    {
                        flexGrid.SetData(i + flexGrid.Rows.Fixed, dt.Columns[j].ColumnName, dt.Rows[i - START_ROW][j]);
                    }
                    else
                    {
                        continue;
                    }
                }
                // 로우에 노드를 만든다
                flexGrid.Rows[i + flexGrid.Rows.Fixed].IsNode = true;
                // 노드의 레벨을 설정(Level)
                flexGrid.Rows[i + flexGrid.Rows.Fixed].Node.Level = ToInt(flexGrid.Rows[i + flexGrid.Rows.Fixed][levelCol]);
            }

            //노드확장 수준 결정
            if (isTopLevel)
                flexGrid.Rows[flexGrid.Rows.Fixed].Node.Collapsed = true;

            for (int i = START_ROW; i < END_ROW; i++)
            {
                //확장레벨보다 하위레벨이면 Collapse처리
                if (ToInt(flexGrid.Rows[i + flexGrid.Rows.Fixed][levelCol]) >= expandLevel)
                    flexGrid.Rows[i + flexGrid.Rows.Fixed].Node.Collapsed = true;
                else
                    flexGrid.Rows[i + flexGrid.Rows.Fixed].Node.Collapsed = false;
            }

            //헤더표시 여부
            if (!isHeader)
            {
                for (int i = 0; i < flexGrid.Rows.Fixed; i++)
                {
                    flexGrid.Rows.Remove(i);
                }

                flexGrid.Rows.Fixed = 0;
            }

            flexGrid.EndUpdate();
            flexGrid.Refresh();
        }

        #endregion

        #region 기타 메서드

        public static int ToInt(object objValue)
        {
            int nullValue = 0;
            return (ToString(objValue) == "" ? nullValue : Convert.ToInt32(objValue));
        }

        public static System.String ToString(object objValue)
        {
            if (objValue == null || objValue == System.DBNull.Value)
                return "";
            return objValue.ToString();
        }

        public static int FlexGridFindCol(C1FlexGrid flexGrid, string colName)
        {
            bool isCract = false;
            //비교값이 같으면 Break;
            for (int i = 0; i < flexGrid.Cols.Count; i++)
            {
                if (flexGrid.Cols[i].Name == colName)
                {
                    isCract = true;
                }

                //비교값이 모두 일치하는 경우
                if (isCract) return i;
            }

            return -1;
        }

        // 파일 확장자 제거
        private static string RemoveExtension(string fileName)
        {
            string ext = Path.GetExtension(fileName);

            if (!string.IsNullOrEmpty(ext))
            {
                int index = fileName.LastIndexOf(".");
                fileName = fileName.Remove(index);
            }

            return fileName;
        }

        #endregion
    }

    /// <summary>
    /// FlexGrid 합계 형식 클래스
    /// </summary>
    public class SubTotalItem
    {
        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="aggregateEnum">합계유형</param>
        /// <param name="level">SubTotal Level(Tree Depth)</param>
        /// <param name="columnName">컬럼명</param>
        /// <param name="groupbyColumnName">그룹 컬럼 명</param>
        /// <param name="caption">그룹 컬럼 명</param>
        public SubTotalItem(AggregateEnum aggregateEnum, int level, string groupbyColumnName, string columnName, string caption)
        {
            this.AggregateEnum = aggregateEnum;
            this.Level = level;
            this.GroupbyColumnName = groupbyColumnName;
            this.ColumnName = columnName;
            this.Caption = caption;
        }

        /// <summary>
        /// 합계 유형
        /// </summary>
        public AggregateEnum AggregateEnum
        {
            get;
            set;
        }

        /// <summary>
        /// SubTotal Level(Tree Depth)
        /// </summary>
        public int Level
        {
            get;
            set;
        }

        /// <summary>
        /// 그룹 설정 컬럼
        /// </summary>
        public string GroupbyColumnName
        {
            get;
            set;
        }

        /// <summary>
        /// 합계 FlexGrid 컬럽
        /// </summary>
        public string ColumnName
        {
            get;
            set;
        }
        /// <summary>
        /// 합계 타이틀
        /// </summary>
        public string Caption
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Grand Total Item
    /// </summary>
    public class GrandTotalItem
    {
        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="aggregateEnum">합계유형</param>
        /// <param name="columnName">컬럼명</param>
        public GrandTotalItem(AggregateEnum aggregateEnum, string columnName)
        {
            this.AggregateEnum = aggregateEnum;
            this.ColumnName = columnName;
        }

        /// <summary>
        /// 합계 유형
        /// </summary>
        public AggregateEnum AggregateEnum
        {
            get;
            set;
        }

        /// <summary>
        /// 합계 FlexGrid 컬럽
        /// </summary>
        public string ColumnName
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Grid Style 형식 지정 
    /// </summary>
    public enum StyleType
    {
        /// <summary>
        /// 기본 값
        /// </summary>
        DefaultCell,
        /// <summary>
        /// 읽기 전용
        /// </summary>
        ReadOnlyCell,
        /// <summary>
        /// Grand Total 값 Cell
        /// </summary>
        GrandTotal,
        /// <summary>
        /// 변경된 Row 스타일
        /// </summary>
        ChangeRow,
        /// <summary>
        /// 새로 추가된 Row 스타일
        /// </summary>
        NewRow,
        /// <summary>
        /// 삭제된 Row 스타일
        /// </summary>
        DeleteRow
    }

    /// <summary>
    /// Grand Total 계산 이벤트 인자값
    /// </summary>
    public class GrandTotalEventArgs : EventArgs
    {
        /// <summary>
        /// GrandTotal 컬럼명
        /// </summary>
        public string ColumnName
        {
            get;
            set;
        }

        /// <summary>
        /// GrandTotal 계산 형식
        /// </summary>
        public AggregateEnum AggregateType
        {
            get;
            set;
        }

        /// <summary>
        /// Total Value
        /// </summary>
        public double TotalValue
        {
            get;
            set;
        }
    }

}
