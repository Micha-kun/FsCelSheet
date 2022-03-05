namespace FsCelSheet

open System
open System.Drawing

[<Struct>]
type ExcelCellSize =
  private
    { ColSpan: int
      RowSpan: int }
  static member MergedCell(colSpan, rowSpan) =
    { ColSpan = colSpan; RowSpan = rowSpan }

  static member SingleCell = ExcelCellSize.MergedCell(1, 1)

  member self.IsMergedCell =
    match self.ColSpan, self.RowSpan with
    | 1, 1 -> false
    | _ -> true

[<AutoOpen>]
module ExcelCellSize =
  let (|ExcelCellSize|) (x: ExcelCellSize) = x.ColSpan, x.RowSpan

  let width (ExcelCellSize (col, _)) = col

  let height (ExcelCellSize (_, row)) = row

[<Struct>]
type ExcelCellPosition =
  private
    { Row: int
      Column: int }
  static member Create(row, column) =
    match row, column with
    | r, _ when r < 1 -> ExcelCellPosition.Create(1, column)
    | _, c when c < 1 -> ExcelCellPosition.Create(row, 1)
    | r, c -> { Row = r; Column = c }

  static member Init = ExcelCellPosition.Create(1, 1)

[<AutoOpen>]
module ExcelCellPosition =
  let (|ExcelCellPosition|) (x: ExcelCellPosition) = x.Row, x.Column

  let move (x, y) (ExcelCellPosition (r, c)) = ExcelCellPosition.Create(x + r, y + c)

  let moveRow x p = move (x, 0) p

  let nextRow p = moveRow 1 p

  let prevRow p = moveRow -1 p

  let moveColumn y p = move (0, y) p

  let nextColumn p = moveColumn 1 p

  let prevColumn p = moveColumn -1 p

  let row (ExcelCellPosition (r, _)) = r

  let column (ExcelCellPosition (_, c)) = c

type ExcelCellDataValue =
  | IntegerValue of int
  | DecimalValue of decimal
  | CurrencyValue of decimal * string
  | PercentageValue of decimal
  | TextValue of string
  | DateTimeValue of DateTime
  | DateValue of DateTime
  | TimeValue of DateTime
  | FormulaValue of string
  static member EmptyValue = TextValue ""

  static member (~-)(v: ExcelCellDataValue) =
    match v with
    | IntegerValue x -> IntegerValue -x
    | DecimalValue x -> DecimalValue -x
    | CurrencyValue (x, s) -> CurrencyValue(-x, s)
    | PercentageValue x ->
      let pText = x.ToString("0.00 %")
      TextValue $"/ {pText}"
    | TextValue x -> TextValue $"/ {x}"
    | DateTimeValue x ->
      let pText = x.ToString("dd/MM/yyyy hh:mm")
      TextValue $" / {pText}"
    | DateValue x ->
      let pText = x.ToString("dd/MM/yyyy")
      TextValue $" / {pText}"
    | TimeValue x ->
      let pText = x.ToString("hh:mm")
      TextValue $" / {pText}"
    | FormulaValue x -> FormulaValue x

  static member (-)(v1: ExcelCellDataValue, v2: ExcelCellDataValue) =
    match v1, v2 with
    | IntegerValue x1, IntegerValue x2 -> IntegerValue(x1 - x2)
    | DecimalValue x1, DecimalValue x2 -> DecimalValue(x1 - x2)
    | CurrencyValue (x1, s1), CurrencyValue (x2, s2) when s1 = s2 -> CurrencyValue(x1 - x2, s1)
    | CurrencyValue (x1, s1), CurrencyValue (x2, s2) ->
      let p1Text = x1.ToString($"""#,##0.00 %s{s1}""")
      let p2Text = x2.ToString($"""#,##0.00 %s{s2}""")
      TextValue $"{p1Text} / {p2Text}"
    | PercentageValue x1, PercentageValue x2 when x1 = x2 -> PercentageValue x1
    | PercentageValue x1, PercentageValue x2 ->
      let p1Text = x1.ToString("0.00 %")
      let p2Text = x2.ToString("0.00 %")
      TextValue $"{p1Text} / {p2Text}"
    | TextValue x1, TextValue x2 when x1 = x2 -> TextValue x1
    | TextValue x1, TextValue x2 -> TextValue $"{x1} / {x2}"
    | DateTimeValue x1, DateTimeValue x2 when x1 = x2 -> DateTimeValue x1
    | DateTimeValue x1, DateTimeValue x2 ->
      let p1Text = x1.ToString("dd/MM/yyyy hh:mm")
      let p2Text = x2.ToString("dd/MM/yyyy hh:mm")
      TextValue $"{p1Text} / {p2Text}"
    | DateValue x1, DateValue x2 when x1.Date = x2.Date -> DateValue x1
    | DateValue x1, DateValue x2 ->
      let p1Text = x1.ToString("dd/MM/yyyy")
      let p2Text = x2.ToString("dd/MM/yyyy")
      TextValue $"{p1Text} / {p2Text}"
    | TimeValue x1, TimeValue x2 when x1.ToString("hh:mm") = x2.ToString("hh:mm") -> TimeValue x1
    | TimeValue x1, TimeValue x2 ->
      let p1Text = x1.ToString("hh:mm")
      let p2Text = x2.ToString("hh:mm")
      TextValue $"{p1Text} / {p2Text}"
    | FormulaValue x1, FormulaValue _ -> FormulaValue x1
    | x, y -> failwithf "Unmatching ExcelCellDataValue: %A - %A" x y

[<Struct>]
type ExcelCellDataFormat =
  | ExcelCellDataFormat of string
  static member Create(format: string) =
    format.Trim().TrimStart('=')
    |> ExcelCellDataFormat

  static member TextFormat = ExcelCellDataFormat.Create "@"
  static member IntegerFormat = ExcelCellDataFormat.Create "#,##0"
  static member DecimalFormat = ExcelCellDataFormat.Create "#,##0.00"
  static member PercentageFormat = ExcelCellDataFormat.Create "0.00 %"
  static member DateTimeFormat = ExcelCellDataFormat.Create "dd/MM/yyyy hh:mm"
  static member DateFormat = ExcelCellDataFormat.Create "dd/MM/yyyy"
  static member TimeFormat = ExcelCellDataFormat.Create "hh:mm"

  static member CurrencyFormat c =
    $"""#,##0.00 %s{c}"""
    |> ExcelCellDataFormat.Create

type ExcelCellData =
  private
    { Value: ExcelCellDataValue
      Format: ExcelCellDataFormat option }
  static member Create(typedData, ?format) = { Value = typedData; Format = format }
  static member IntegerCell x = ExcelCellData.Create(IntegerValue x)

  static member DecimalCell x =
    ExcelCellData.Create(DecimalValue x, ExcelCellDataFormat.DecimalFormat)

  static member CurrencyCell(x, c) =
    ExcelCellData.Create(CurrencyValue(x, c), ExcelCellDataFormat.CurrencyFormat c)

  static member PercentageCell x =
    ExcelCellData.Create(PercentageValue x, ExcelCellDataFormat.PercentageFormat)

  static member TextCell x = ExcelCellData.Create(TextValue x)

  static member DateTimeCell x =
    ExcelCellData.Create(DateTimeValue x, ExcelCellDataFormat.DateTimeFormat)

  static member DateCell x =
    ExcelCellData.Create(DateValue x, ExcelCellDataFormat.DateFormat)

  static member TimeCell x =
    ExcelCellData.Create(TimeValue x, ExcelCellDataFormat.TimeFormat)

  static member FormulaCell x = ExcelCellData.Create(FormulaValue x)
  member self.WithFormat f = ExcelCellData.Create(self.Value, f)

type ExcelCellDataValue with
  member this.CellData =
    match this with
    | IntegerValue x -> ExcelCellData.IntegerCell x
    | DecimalValue x -> ExcelCellData.DecimalCell x
    | CurrencyValue (x, c) -> ExcelCellData.CurrencyCell(x, c)
    | PercentageValue x -> ExcelCellData.PercentageCell x
    | TextValue x -> ExcelCellData.TextCell x
    | DateTimeValue x -> ExcelCellData.DateTimeCell x
    | DateValue x -> ExcelCellData.DateCell x
    | TimeValue x -> ExcelCellData.TimeCell x
    | FormulaValue x -> ExcelCellData.FormulaCell x

[<AutoOpen>]
module ExcelCellData =
  let (|Simple|Formatted|) (x: ExcelCellData) =
    match x.Format with
    | Some format -> Formatted(x.Value, format)
    | None -> Simple x.Value

[<Struct>]
type ExcelCellBorder =
  | NoBorder
  | Solid of Color

[<Struct>]
type ExcelCellVerticalAlignement =
  | Top
  | Middle
  | Bottom

[<Struct>]
type ExcelCellHorizontalAlignement =
  | Left
  | Center
  | Right

type ExcelCellStyleProperty =
  | Bold of bool
  | Italic of bool
  | Underline of bool
  | WrapText of bool
  | Border of ExcelCellBorder
  | VerticalAlignement of ExcelCellVerticalAlignement
  | HorizontalAlignement of ExcelCellHorizontalAlignement
  | BackgroundColor of Color
  | ForegroundColor of Color
  | FontSize of int
  | FontFamily of string
  member this.Key =
    match this with
    | Bold _ -> "Bold"
    | Italic _ -> "Italic"
    | Underline _ -> "Underline"
    | WrapText _ -> "WrapText"
    | Border _ -> "Border"
    | VerticalAlignement _ -> "VerticalAlignement"
    | HorizontalAlignement _ -> "HorizontalAlignement"
    | BackgroundColor _ -> "BackgroundColor"
    | ForegroundColor _ -> "ForegroundColor"
    | FontSize _ -> "FontSize"
    | FontFamily _ -> "FontFamily"

type ExcelCellStyle = ExcelCellStyleProperty list

module ExcelCellStyle =
  let emptyStyle: ExcelCellStyle = []

  let optimize (list: ExcelCellStyle) : ExcelCellStyle =
    List.foldBack
      (fun (item: ExcelCellStyleProperty) coll ->
        match Map.tryFind item.Key coll with
        | Some _ -> Map.change item.Key (fun _ -> Some item) coll
        | None -> Map.add item.Key item coll)
      list
      Map.empty
    |> Map.toList
    |> List.map snd

type ExcelCell =
  { Data: ExcelCellData
    Position: ExcelCellPosition
    Size: ExcelCellSize
    Style: ExcelCellStyle }
