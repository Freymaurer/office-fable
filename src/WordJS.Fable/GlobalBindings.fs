namespace OfficeJS.Fable

open Fable.Core
open Fable.Core.JsInterop


module Office =

    open OfficeJS.Fable

    [<Global>]
    let Office : Office.IExports = jsNative

module Word =

    [<Global>]
    let Word : Word.IExports = jsNative

    [<Global>]
    let WordRangeLoadOptions : Word.Interfaces.RangeLoadOptions = jsNative