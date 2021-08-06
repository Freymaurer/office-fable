module OfficeInterop

open Fable.Core
open Fable.Core.JsInterop

open OfficeJS
open Excel
open Word

[<Global>]
let Office : Office.IExports = jsNative

[<Global>]
//[<CompiledName("Office.Excel")>]
let Excel : Excel.IExports = jsNative

[<Global>]
let Word : Word.IExports = jsNative

[<Global>]
let OneNote : OneNote.IExports = jsNative

[<Global>]
let PowerPoint : PowerPoint.IExports = jsNative

[<Global>]
let Visio : Visio.IExports = jsNative

[<Global>]
let RangeLoadOptions : Interfaces.RangeLoadOptions = jsNative

