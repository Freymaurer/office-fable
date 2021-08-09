namespace WordJS.Fable

open Fable.Core
open Fable.Core.JsInterop


module GlobalBindings =

    open WordJS.Fable

    [<Global>]
    let Office : Office.IExports = jsNative


    [<Global>]
    let Word : Word.IExports = jsNative

    [<Global>]
    let WordRangeLoadOptions : Word.Interfaces.RangeLoadOptions = jsNative