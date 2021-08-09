namespace OfficeJS.Fable

open System
open Fable.Core
open Fable.Core.JS
open Browser.Types

module rec PowerPoint =

    type [<AllowNullLiteral>] IExports =
        abstract Application: ApplicationStatic
        abstract Presentation: PresentationStatic
        abstract Slide: SlideStatic
        abstract SlideCollection: SlideCollectionStatic
        abstract RequestContext: RequestContextStatic
        /// <summary>Executes a batch script that performs actions on the PowerPoint object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</summary>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.</param>
        abstract run: batch: (PowerPoint.RequestContext -> OfficeExtension.IPromise<'T>) -> OfficeExtension.IPromise<'T>
        /// <summary>Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</summary>
        /// <param name="object">- A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".</param>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.</param>
        abstract run: ``object``: OfficeExtension.ClientObject * batch: (PowerPoint.RequestContext -> OfficeExtension.IPromise<'T>) -> OfficeExtension.IPromise<'T>
        /// <summary>Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.</summary>
        /// <param name="objects">- An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".</param>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.</param>
        abstract run: objects: ResizeArray<OfficeExtension.ClientObject> * batch: (PowerPoint.RequestContext -> OfficeExtension.IPromise<'T>) -> OfficeExtension.IPromise<'T>
        /// <summary>Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a base64-encoded .pptx file.
        /// 
        /// [Api set: PowerPointApi 1.1]</summary>
        /// <param name="base64File">Optional. The base64-encoded .pptx file. The default value is null.</param>
        abstract createPresentation: ?base64File: string -> Promise<unit>

    /// [Api set: PowerPointApi 1.0]
    type [<AllowNullLiteral>] Application =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original PowerPoint.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> ApplicationToJSONReturn

    type [<AllowNullLiteral>] ApplicationToJSONReturn =
        [<Emit "$0[$1]{{=$2}}">] abstract Item: key: string -> string with get, set

    /// [Api set: PowerPointApi 1.0]
    type [<AllowNullLiteral>] ApplicationStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Application
        /// Create a new instance of PowerPoint.Application object
        abstract newObject: context: OfficeExtension.ClientRequestContext -> PowerPoint.Application

    /// [Api set: PowerPointApi 1.0]
    type [<AllowNullLiteral>] Presentation =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Returns an ordered collection of slides in the presentation.
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract slides: PowerPoint.SlideCollection
        abstract title: string
        /// <summary>Inserts the specified slides from a presentation into the current presentation.
        /// 
        /// [Api set: PowerPointApi 1.2]</summary>
        /// <param name="base64File">The base64-encoded string representing the source presentation file.</param>
        /// <param name="options">The options that define which slides will be inserted, where the new slides will go, and which presentation's formatting will be used.</param>
        abstract insertSlidesFromBase64: base64File: string * ?options: PowerPoint.InsertSlideOptions -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: PowerPoint.Interfaces.PresentationLoadOptions -> PowerPoint.Presentation
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> PowerPoint.Presentation
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: PresentationLoadPropertyNamesAndPaths -> PowerPoint.Presentation
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original PowerPoint.Presentation object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> PowerPoint.Interfaces.PresentationData

    type [<AllowNullLiteral>] PresentationLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// [Api set: PowerPointApi 1.0]
    type [<AllowNullLiteral>] PresentationStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Presentation

    type [<StringEnum>] [<RequireQualifiedAccess>] InsertSlideFormatting =
        | [<CompiledName "KeepSourceFormatting">] KeepSourceFormatting
        | [<CompiledName "UseDestinationTheme">] UseDestinationTheme

    /// Represents the available options when inserting slides.
    /// 
    /// [Api set: PowerPointApi 1.2]
    type [<AllowNullLiteral>] InsertSlideOptions =
        /// Specifies which formatting to use during slide insertion.
        ///           The default option is to use "KeepSourceFormatting".
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract formatting: U2<PowerPoint.InsertSlideFormatting, string> option with get, set
        /// Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.
        ///           The order of these slides is preserved during the insertion.
        ///           If any of the source slides are not found, or if the IDs are invalid, the operation throws a `SlideNotFound` exception and no slides will be inserted.
        ///           All of the source slides will be inserted when `sourceSlideIds` is not provided (this is the default behavior).
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract sourceSlideIds: ResizeArray<string> option with get, set
        /// Specifies where in the presentation the new slides will be inserted. The new slides will be inserted after the slide with the given slide ID.
        ///           If `targetSlideId` is not provided, the slides will be inserted at the beginning of the presentation.
        ///           If `targetSlideId` is invalid or if it is pointing to a non-existing slide, the operation throws a `SlideNotFound` exception and no slides will be inserted.
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract targetSlideId: string option with get, set

    /// Represents a single slide of a presentation.
    /// 
    /// [Api set: PowerPointApi 1.2]
    type [<AllowNullLiteral>] Slide =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the unique ID of the slide.
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract id: string
        /// Deletes the slide from the presentation. Does nothing if the slide does not exist.
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract delete: unit -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: PowerPoint.Interfaces.SlideLoadOptions -> PowerPoint.Slide
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> PowerPoint.Slide
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: SlideLoadPropertyNamesAndPaths -> PowerPoint.Slide
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original PowerPoint.Slide object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> PowerPoint.Interfaces.SlideData

    type [<AllowNullLiteral>] SlideLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a single slide of a presentation.
    /// 
    /// [Api set: PowerPointApi 1.2]
    type [<AllowNullLiteral>] SlideStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Slide

    /// Represents the collection of slides in the presentation.
    /// 
    /// [Api set: PowerPointApi 1.2]
    type [<AllowNullLiteral>] SlideCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<PowerPoint.Slide>
        /// Gets the number of slides in the collection.
        /// 
        /// [Api set: PowerPointApi 1.2]
        abstract getCount: unit -> OfficeExtension.ClientResult<float>
        /// <summary>Gets a slide using its unique ID.
        /// 
        /// [Api set: PowerPointApi 1.2]</summary>
        /// <param name="key">The ID of the slide.</param>
        abstract getItem: key: string -> PowerPoint.Slide
        /// <summary>Gets a slide using its zero-based index in the collection. Slides are stored in the same order as they
        ///           are shown in the presentation.
        /// 
        /// [Api set: PowerPointApi 1.2]</summary>
        /// <param name="index">The index of the slide in the collection.</param>
        abstract getItemAt: index: float -> PowerPoint.Slide
        /// <summary>Gets a slide using its unique ID. If such a slide does not exist, an object with an `isNullObject` property set to true is returned. For further information,
        ///           see {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods
        ///           and properties}.
        /// 
        /// [Api set: PowerPointApi 1.2]</summary>
        /// <param name="id">The ID of the slide.</param>
        abstract getItemOrNullObject: id: string -> PowerPoint.Slide
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> PowerPoint.SlideCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> PowerPoint.SlideCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> PowerPoint.SlideCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `PowerPoint.SlideCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> PowerPoint.Interfaces.SlideCollectionData

    /// Represents the collection of slides in the presentation.
    /// 
    /// [Api set: PowerPointApi 1.2]
    type [<AllowNullLiteral>] SlideCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SlideCollection

    type [<StringEnum>] [<RequireQualifiedAccess>] ErrorCodes =
        | [<CompiledName "GeneralException">] GeneralException

    module Interfaces =

        /// Provides ways to load properties of only a subset of members of a collection.
        type [<AllowNullLiteral>] CollectionLoadOptions =
            /// Specify the number of items in the queried collection to be included in the result.
            abstract ``$top``: float option with get, set
            /// Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            abstract ``$skip``: float option with get, set

        /// An interface for updating data on the SlideCollection object, for use in `slideCollection.set({ ... })`.
        type [<AllowNullLiteral>] SlideCollectionUpdateData =
            abstract items: ResizeArray<PowerPoint.Interfaces.SlideData> option with get, set

        /// An interface describing the data returned by calling `presentation.toJSON()`.
        type [<AllowNullLiteral>] PresentationData =
            abstract title: string option with get, set

        /// An interface describing the data returned by calling `slide.toJSON()`.
        type [<AllowNullLiteral>] SlideData =
            /// Gets the unique ID of the slide.
            /// 
            /// [Api set: PowerPointApi 1.2]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling `slideCollection.toJSON()`.
        type [<AllowNullLiteral>] SlideCollectionData =
            abstract items: ResizeArray<PowerPoint.Interfaces.SlideData> option with get, set

        /// [Api set: PowerPointApi 1.0]
        type [<AllowNullLiteral>] PresentationLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            abstract title: bool option with get, set

        /// Represents a single slide of a presentation.
        /// 
        /// [Api set: PowerPointApi 1.2]
        type [<AllowNullLiteral>] SlideLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the unique ID of the slide.
            /// 
            /// [Api set: PowerPointApi 1.2]
            abstract id: bool option with get, set

        /// Represents the collection of slides in the presentation.
        /// 
        /// [Api set: PowerPointApi 1.2]
        type [<AllowNullLiteral>] SlideCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the unique ID of the slide.
            /// 
            /// [Api set: PowerPointApi 1.2]
            abstract id: bool option with get, set

    /// The RequestContext object facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the request context is required to get access to the PowerPoint object model from the add-in.
    type [<AllowNullLiteral>] RequestContext =
        inherit OfficeCore.RequestContext
        abstract presentation: Presentation
        abstract application: Application

    /// The RequestContext object facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the request context is required to get access to the PowerPoint object model from the add-in.
    type [<AllowNullLiteral>] RequestContextStatic =
        [<Emit "new $0($1...)">] abstract Create: ?url: string -> RequestContext