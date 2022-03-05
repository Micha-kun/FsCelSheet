namespace FsCelSheet

open FSharpx

/// ExcelComponentCell
type ExcelComponentCell =
  private
    { Data: ExcelCellData
      Position: ExcelCellPosition
      Size: ExcelCellSize
      Style: ExcelCellStyle option }

  static member Create(data, pos, size, ?styles) =
    { Data = data
      Position = pos
      Size = size
      Style = styles }

  member self.WithFormat f =
    { self with Data = self.Data.WithFormat f }

  member self.WithStyle(s: ExcelCellStyle) =
    { self with
        Style =
          self.Style
          |> Option.option s (fun style -> style @ s)
          |> Some }

  member self.AsMerged(colSpan, rowSpan) =
    { self with Size = ExcelCellSize.MergedCell(colSpan, rowSpan) }

[<AutoOpen>]
module ExcelComponentCell =
  let (|ExcelComponentCell|) (x: ExcelComponentCell) = (x.Data, x.Position, x.Size, x.Style)

/// ExcelComponent
type ExcelComponent =
  private
  | ComponentCell of ExcelComponentCell
  | ComponentBlock of ExcelComponent list * ExcelCellStyle option
  static member CreateCell cell = ComponentCell cell
  static member CreateBlock(components, ?style) = ComponentBlock(components, style)

[<AutoOpen>]
module ExcelComponent =
  let (|ComponentCell|ComponentBlock|) (x: ExcelComponent) =
    match x with
    | ComponentCell cell -> ComponentCell cell
    | ComponentBlock (components, style) -> ComponentBlock(components, style)

  let fromCellData cellData position =
    ExcelComponentCell.Create(cellData, position, ExcelCellSize.SingleCell)
    |> ExcelComponent.CreateCell

  // Decorate functions
  let withStyle s componentCell =
    match componentCell with
    | ComponentCell cell -> cell.WithStyle s |> ExcelComponent.CreateCell
    | ComponentBlock (components, style) ->
      ComponentBlock(
        components,
        style
        |> Option.option s (fun style -> style @ s)
        |> Some
      )

  let withDefaultStyle s componentCell =
    match componentCell with
    | ComponentCell cell -> cell.WithStyle s |> ExcelComponent.CreateCell
    | ComponentBlock (components, style) ->
      ComponentBlock(
        components,
        style
        |> Option.option s (fun style -> s @ style)
        |> Some
      )

  let rec withFormat f componentCell =
    match componentCell with
    | ComponentCell cell -> cell.WithFormat f |> ExcelComponent.CreateCell
    | ComponentBlock (components, style) ->
      let formattedComponents = components |> List.map (withFormat f)
      ComponentBlock(formattedComponents, style)

  let rec asMerged (colSpan, rowSpan) componentCell =
    match componentCell with
    | ComponentCell c -> c.AsMerged(colSpan, rowSpan) |> ComponentCell
    | ComponentBlock (items, style) -> ComponentBlock(items |> List.map (asMerged (colSpan, rowSpan)), style)

  // Movement functions
  let rec nextRow componentCell =
    match componentCell with
    | ComponentCell cell -> cell.Position |> moveRow (height cell.Size)
    | ComponentBlock (components, _) ->
      let cells = components |> List.map nextRow

      let (ExcelCellPosition (maxRow, _)) =
        cells
        |> List.maxBy (fun (ExcelCellPosition (r, _)) -> r)

      let (ExcelCellPosition (_, minCol)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (_, c)) -> c)

      ExcelCellPosition.Create(maxRow, minCol)

  let rec prevRow componentCell =
    match componentCell with
    | ComponentCell cell -> cell.Position |> ExcelCellPosition.prevRow
    | ComponentBlock (components, _) ->
      let cells = components |> List.map prevRow

      let (ExcelCellPosition (minRow, _)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (r, _)) -> r)

      let (ExcelCellPosition (_, minCol)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (_, c)) -> c)

      ExcelCellPosition.Create(minRow, minCol)

  let rec nextColumn componentCell =
    match componentCell with
    | ComponentCell cell -> cell.Position |> moveColumn (width cell.Size)
    | ComponentBlock (components, _) ->
      let cells = components |> List.map nextColumn

      let (ExcelCellPosition (minRow, _)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (r, _)) -> r)

      let (ExcelCellPosition (_, maxCol)) =
        cells
        |> List.maxBy (fun (ExcelCellPosition (_, c)) -> c)

      ExcelCellPosition.Create(minRow, maxCol)

  let rec prevColumn componentCell =
    match componentCell with
    | ComponentCell cell -> cell.Position |> ExcelCellPosition.prevColumn
    | ComponentBlock (components, _) ->
      let cells = components |> List.map prevColumn

      let (ExcelCellPosition (minRow, _)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (r, _)) -> r)

      let (ExcelCellPosition (_, minCol)) =
        cells
        |> List.minBy (fun (ExcelCellPosition (_, c)) -> c)

      ExcelCellPosition.Create(minRow, minCol)

  // Conversor
  let rec toExcelCellList defaultStyle controlCell : ExcelCell list =
    match withDefaultStyle defaultStyle controlCell with
    | ComponentCell (ExcelComponentCell (data, position, size, Some style)) ->
      { ExcelCell.Data = data
        Position = position
        Size = size
        Style = style }
      |> List.singleton
    | ComponentCell (ExcelComponentCell (data, position, size, None)) ->
      { ExcelCell.Data = data
        Position = position
        Size = size
        Style = defaultStyle }
      |> List.singleton
    | ComponentBlock (items, Some style) -> items |> List.collect (toExcelCellList style)
    | ComponentBlock (items, None) ->
      items
      |> List.collect (toExcelCellList defaultStyle)

/// ExcelComponentFactory
type ExcelComponentFactory = ExcelCellPosition -> ExcelComponent

[<AutoOpen>]
module ExcelComponentFactory =
  // Excel Cell functions
  let text = ExcelCellData.TextCell

  let textCell: _ -> ExcelComponentFactory = text >> fromCellData

  let empty = text ""

  let emptyCell = textCell ""

  let textCellOpt textOpt =
    match textOpt with
    | Some txt -> textCell txt
    | None -> emptyCell

  let time = ExcelCellData.TimeCell

  let timeCell: _ -> ExcelComponentFactory = time >> fromCellData

  let timeCellOpt timeOpt =
    match timeOpt with
    | Some v -> timeCell v
    | None -> emptyCell

  let currency = ExcelCellData.CurrencyCell

  let currencyCell: _ -> ExcelComponentFactory = currency >> fromCellData

  let currencyCellOpt currencyOpt =
    match currencyOpt with
    | Some v -> currencyCell v
    | None -> emptyCell

  let dateTime = ExcelCellData.DateTimeCell

  let dateTimeCell: _ -> ExcelComponentFactory = dateTime >> fromCellData

  let dateTimeCellOpt dateTimeOpt =
    match dateTimeOpt with
    | Some v -> dateTimeCell v
    | None -> emptyCell

  let date = ExcelCellData.DateCell

  let dateCell: _ -> ExcelComponentFactory = date >> fromCellData

  let dateCellOpt dateOpt =
    match dateOpt with
    | Some v -> dateCell v
    | None -> emptyCell

  let percentage = ExcelCellData.PercentageCell

  let percentageCell: _ -> ExcelComponentFactory = percentage >> fromCellData

  let percentageCellOpt percentageOpt =
    match percentageOpt with
    | Some v -> percentageCell v
    | None -> emptyCell

  let integer = ExcelCellData.IntegerCell

  let integerCell: _ -> ExcelComponentFactory = integer >> fromCellData

  let integerCellOpt integerOpt =
    match integerOpt with
    | Some v -> integerCell v
    | None -> emptyCell

  let decimal = ExcelCellData.DecimalCell

  let decimalCell: _ -> ExcelComponentFactory = decimal >> fromCellData

  let decimalCellOpt decimalOpt =
    match decimalOpt with
    | Some v -> decimalCell v
    | None -> emptyCell

  let formula = ExcelCellData.FormulaCell

  let formulaCell: _ -> ExcelComponentFactory = formula >> fromCellData

  // Decoration functions
  let asMerged size (cellFunc: ExcelComponentFactory) : ExcelComponentFactory = cellFunc >> asMerged size

  let withFormat format (cellFunc: ExcelComponentFactory) : ExcelComponentFactory = cellFunc >> withFormat format

  let withSimpleFormat format (cellFunc: ExcelComponentFactory) : ExcelComponentFactory =
    let f = ExcelCellDataFormat.Create format
    withFormat f cellFunc

  let withStyleOpt style (cellFunc: ExcelComponentFactory) : ExcelComponentFactory =
    match style with
    | Some s -> cellFunc >> withStyle s
    | None -> cellFunc

  let withDefaultStyleOpt defaultStyle (cellFunc: ExcelComponentFactory) : ExcelComponentFactory =
    match defaultStyle with
    | Some s -> cellFunc >> withDefaultStyle s
    | None -> cellFunc

  let withStyle style (cellFunc: ExcelComponentFactory) : ExcelComponentFactory = cellFunc >> withStyle style

  let withDefaultStyle defaultStyle (cellFunc: ExcelComponentFactory) : ExcelComponentFactory =
    cellFunc >> withDefaultStyle defaultStyle

  let setBold = Bold >> List.singleton >> withStyle

  let setItalic = Italic >> List.singleton >> withStyle

  let setUnderline = Underline >> List.singleton >> withStyle

  let setWrapText = WrapText >> List.singleton >> withStyle

  let setBorder = Border >> List.singleton >> withStyle

  let setVerticalAlignement = VerticalAlignement >> List.singleton >> withStyle

  let setHorizontalAlignement =
    HorizontalAlignement
    >> List.singleton
    >> withStyle

  let setBackgroundColor = BackgroundColor >> List.singleton >> withStyle

  let setForegroundColor = ForegroundColor >> List.singleton >> withStyle

  let setFontSize = FontSize >> List.singleton >> withStyle

  let setFontFamily = FontFamily >> List.singleton >> withStyle

  // Composition functions
  let concat nextPosFunc (factoryList: ExcelComponentFactory list) =
    if factoryList |> List.isEmpty then
      failwith "Error vacío"
    else
      let cellBlockCreator: ExcelComponentFactory =
        fun pos ->
          let cells =
            factoryList
            |> List.mapFold
                 (fun p componentFunc ->
                   let cell = componentFunc p
                   cell, nextPosFunc cell)
                 pos
            |> fst

          ExcelComponent.CreateBlock cells

      cellBlockCreator

  let concatAsColumn cellFuncList : ExcelComponentFactory = concat nextRow cellFuncList

  let tryConcatAsColumn cellFuncList : ExcelComponentFactory option =
    match cellFuncList with
    | [] -> None
    | x -> Some <| concat nextRow x

  let concatAsRow cellFuncList : ExcelComponentFactory = concat nextColumn cellFuncList

  let tryConcatAsRow cellFuncList : ExcelComponentFactory option =
    match cellFuncList with
    | [] -> None
    | x -> Some <| concat nextColumn x

  // Movement functions
  let private moveCtor func x (creator: ExcelComponentFactory) : ExcelComponentFactory =
    let movedCreator: ExcelComponentFactory =
      fun pos ->
        let newPos = pos |> func x
        creator newPos

    movedCreator

  let moveRow r creator = moveCtor moveRow r creator

  let nextRow creator = moveRow 1 creator

  let prevRow creator = moveRow -1 creator

  let moveColumn c creator = moveCtor moveColumn c creator

  let nextColumn creator = moveColumn 1 creator

  let prevColumn creator = moveColumn -1 creator

  let move (r, c) creator = moveRow r creator |> moveColumn c
