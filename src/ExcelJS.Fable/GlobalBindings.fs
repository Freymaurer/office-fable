namespace OfficeJS.Fable

open Fable.Core
open Fable.Core.JsInterop

module GlobalBindings =

    open OfficeJS.Fable

    [<Global>]
    let Office : Office.IExports = jsNative


    [<Global>]
    //[<CompiledName("Office.Excel")>]
    let Excel : Excel.IExports = jsNative

    [<Global>]
    let ExcelRangeLoadOptions : Excel.Interfaces.RangeLoadOptions = jsNative