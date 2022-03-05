namespace FsCelSheet.EPPlus4

open OfficeOpenXml

open FsCelSheet

[<AutoOpen>]
module SheetRender =
  let toWorksheet (worksheet: ExcelWorksheet) factory =
    factory ExcelCellPosition.Init
    |> toExcelCellList ExcelCellStyle.Empty
    |> List.iter (renderCell worksheet >> ignore)

    worksheet
