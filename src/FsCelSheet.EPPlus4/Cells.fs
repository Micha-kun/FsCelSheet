namespace FsCelSheet.EPPlus4

open OfficeOpenXml
open OfficeOpenXml.Style

open FsCelSheet

[<AutoOpen>]
module Cells =
  let setCellFormat (ExcelCellDataFormat format) (cellRange: ExcelRange) =
    cellRange.Style.Numberformat.Format <- format
    cellRange

  let rec renderCellDataValue cellDataValue (cellRange: ExcelRange) =
    match cellDataValue with
    | IntegerValue x -> cellRange.Value <- x
    | CurrencyValue(x, _) -> cellRange.Value <- x
    | AccountingValue(x, _) -> cellRange.Value <- x
    | DecimalValue x -> cellRange.Value <- x
    | PercentageValue x -> cellRange.Value <- x
    | TextValue x -> cellRange.Value <- x
    | DateTimeValue x -> cellRange.Value <- x
    | DateValue x -> cellRange.Value <- x
    | TimeValue x -> cellRange.Value <- x
    | FormulaValue x -> cellRange.Formula <- x
    | OptionalValue x ->
      match x with
      | Some v ->
        renderCellDataValue v (cellRange: ExcelRange)
        |> ignore
      | None -> cellRange.Value <- ""

    cellRange

  let renderCellData cellData (cellRange: ExcelRange) =
    match cellData with
    | Simple x -> renderCellDataValue x cellRange
    | Formatted(x, f) -> renderCellDataValue x cellRange |> setCellFormat f

  let private renderBold (cellRange: ExcelRange) property =
    cellRange.Style.Font.Bold <- property
    cellRange

  let private renderItalic (cellRange: ExcelRange) property =
    cellRange.Style.Font.Italic <- property
    cellRange

  let private renderUnderline (cellRange: ExcelRange) property =
    cellRange.Style.Font.UnderLineType <- ExcelUnderLineType.Single
    cellRange.Style.Font.UnderLine <- property
    cellRange

  let private renderWrapText (cellRange: ExcelRange) property =
    cellRange.Style.WrapText <- property
    cellRange

  let private renderBorder (cellRange: ExcelRange) property =
    match property with
    | NoBorder -> ()
    | Solid color -> cellRange.Style.Border.BorderAround(ExcelBorderStyle.Thin, color)

    cellRange

  let private renderVerticalAlignement (cellRange: ExcelRange) property =
    match property with
    | Top -> cellRange.Style.VerticalAlignment <- ExcelVerticalAlignment.Top
    | Middle -> cellRange.Style.VerticalAlignment <- ExcelVerticalAlignment.Center
    | Bottom -> cellRange.Style.VerticalAlignment <- ExcelVerticalAlignment.Bottom

    cellRange

  let private renderHorizontalAlignement (cellRange: ExcelRange) property =
    match property with
    | Left -> cellRange.Style.HorizontalAlignment <- ExcelHorizontalAlignment.Left
    | Center -> cellRange.Style.HorizontalAlignment <- ExcelHorizontalAlignment.Center
    | Right -> cellRange.Style.HorizontalAlignment <- ExcelHorizontalAlignment.Right

    cellRange

  let private renderBackgroundColor (cellRange: ExcelRange) property =
    cellRange.Style.Fill.PatternType <- ExcelFillStyle.Solid
    cellRange.Style.Fill.BackgroundColor.SetColor property
    cellRange

  let private renderForegroundColor (cellRange: ExcelRange) property =
    cellRange.Style.Font.Color.SetColor property
    cellRange

  let private renderFontSize (cellRange: ExcelRange) property =
    cellRange.Style.Font.Size <- (property |> float32)
    cellRange

  let private renderFontFamily (cellRange: ExcelRange) property =
    cellRange.Style.Font.Name <- property
    cellRange

  let private renderCellStyleProperty property (cellRange: ExcelRange) =
    match property with
    | Bold x -> renderBold cellRange x
    | Italic x -> renderItalic cellRange x
    | Underline x -> renderUnderline cellRange x
    | WrapText x -> renderWrapText cellRange x
    | Border x -> renderBorder cellRange x
    | VerticalAlignement x -> renderVerticalAlignement cellRange x
    | HorizontalAlignement x -> renderHorizontalAlignement cellRange x
    | BackgroundColor x -> renderBackgroundColor cellRange x
    | ForegroundColor x -> renderForegroundColor cellRange x
    | FontSize x -> renderFontSize cellRange x
    | FontFamily x -> renderFontFamily cellRange x

  let renderCellStyle (style: ExcelCellStyle) (cellRange: ExcelRange) =
    List.foldBack renderCellStyleProperty (style |> ExcelCellStyle.optimize) cellRange

  let renderCell (worksheet: ExcelWorksheet) (cell: ExcelCell) =
    let ExcelCellSize(colSpan, rowSpan), ExcelCellPosition(row, column) =
      cell.Size, cell.Position

    let cellRange =
      if cell.Size.IsMergedCell then
        let range = worksheet.Cells.[row, column, row + rowSpan - 1, column + colSpan - 1]

        range.Merge <- true
        range
      else
        worksheet.Cells.[row, column]

    cellRange
    |> renderCellData cell.Data
    |> renderCellStyle cell.Style
