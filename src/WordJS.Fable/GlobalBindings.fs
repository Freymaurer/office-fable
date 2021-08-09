namespace OfficeJS.Fable

open Fable.Core
open Fable.Core.JsInterop


module GlobalBindings =

    open OfficeJS.Fable

    [<Global>]
    let Office : Office.IExports = jsNative


    [<Global>]
    let Word : Word.IExports = jsNative

    [<Global>]
    let WordRangeLoadOptions : Word.Interfaces.RangeLoadOptions = jsNative