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
    

    [<Global>]
    let Word : Word.IExports = jsNative

    [<Global>]
    let WordRangeLoadOptions : Word.Interfaces.RangeLoadOptions = jsNative

    
    [<Global>]
    let OneNote : OneNote.IExports = jsNative


    [<Global>]
    let PowerPoint : PowerPoint.IExports = jsNative
    
    
    [<Global>]
    let Visio : Visio.IExports = jsNative