namespace OfficeJS.Fable

open Fable.Core
open Fable.Core.JsInterop


module Office =

    open OfficeJS.Fable

    [<Global>]
    let Office : Office.IExports = jsNative


module Excel =

    [<Global>]
    //[<CompiledName("Office.Excel")>]
    let Excel : Excel.IExports = jsNative

    [<Global>]
    let ExcelRangeLoadOptions : Excel.Interfaces.RangeLoadOptions = jsNative
    

module Word =

    [<Global>]
    let Word : Word.IExports = jsNative

    [<Global>]
    let WordRangeLoadOptions : Word.Interfaces.RangeLoadOptions = jsNative


module OneNote =
    
    [<Global>]
    let OneNote : OneNote.IExports = jsNative


module PowerPoint =

    [<Global>]
    let PowerPoint : PowerPoint.IExports = jsNative
    

module Visio =
    
    [<Global>]
    let Visio : Visio.IExports = jsNative