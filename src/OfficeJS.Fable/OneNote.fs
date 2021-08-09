namespace OfficeJS.Fable

open System
open Fable.Core
open Fable.Core.JS
open Browser.Types


module rec OneNote =

    type [<AllowNullLiteral>] IExports =
        abstract Application: ApplicationStatic
        abstract InkAnalysis: InkAnalysisStatic
        abstract InkAnalysisParagraph: InkAnalysisParagraphStatic
        abstract InkAnalysisParagraphCollection: InkAnalysisParagraphCollectionStatic
        abstract InkAnalysisLine: InkAnalysisLineStatic
        abstract InkAnalysisLineCollection: InkAnalysisLineCollectionStatic
        abstract InkAnalysisWord: InkAnalysisWordStatic
        abstract InkAnalysisWordCollection: InkAnalysisWordCollectionStatic
        abstract FloatingInk: FloatingInkStatic
        abstract InkStroke: InkStrokeStatic
        abstract InkStrokeCollection: InkStrokeCollectionStatic
        abstract InkWord: InkWordStatic
        abstract InkWordCollection: InkWordCollectionStatic
        abstract Notebook: NotebookStatic
        abstract NotebookCollection: NotebookCollectionStatic
        abstract SectionGroup: SectionGroupStatic
        abstract SectionGroupCollection: SectionGroupCollectionStatic
        abstract Section: SectionStatic
        abstract SectionCollection: SectionCollectionStatic
        abstract Page: PageStatic
        abstract PageCollection: PageCollectionStatic
        abstract PageContent: PageContentStatic
        abstract PageContentCollection: PageContentCollectionStatic
        abstract Outline: OutlineStatic
        abstract Paragraph: ParagraphStatic
        abstract ParagraphCollection: ParagraphCollectionStatic
        abstract NoteTag: NoteTagStatic
        abstract RichText: RichTextStatic
        abstract Image: ImageStatic
        abstract Table: TableStatic
        abstract TableRow: TableRowStatic
        abstract TableRowCollection: TableRowCollectionStatic
        abstract TableCell: TableCellStatic
        abstract TableCellCollection: TableCellCollectionStatic
        abstract RequestContext: RequestContextStatic
        /// <summary>Executes a batch script that performs actions on the OneNote object model, using a new request context. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</summary>
        /// <param name="batch">- A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the OneNote application. Since the Office add-in and the OneNote application run in two different processes, the request context is required to get access to the OneNote object model from the add-in.</param>
        abstract run: batch: (OneNote.RequestContext -> Promise<'T>) -> Promise<'T>
        /// <summary>Executes a batch script that performs actions on the OneNote object model, using the request context of a previously-created API object.</summary>
        /// <param name="object">- A previously-created API object. The batch will use the same request context as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".</param>
        /// <param name="batch">- A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</param>
        abstract run: ``object``: OfficeExtension.ClientObject * batch: (OneNote.RequestContext -> Promise<'T>) -> Promise<'T>
        /// <summary>Executes a batch script that performs actions on the OneNote object model, using the request context of previously-created API objects.</summary>
        /// <param name="batch">- A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</param>
        abstract run: objects: ResizeArray<OfficeExtension.ClientObject> * batch: (OneNote.RequestContext -> Promise<'T>) -> Promise<'T>

    /// Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Application =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of notebooks that are open in the OneNote application instance. In OneNote on the web, only one notebook at a time is open in the application instance. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract notebooks: OneNote.NotebookCollection
        /// Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveNotebook: unit -> OneNote.Notebook
        /// Gets the active notebook if one exists. If no notebook is active, returns null.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveNotebookOrNull: unit -> OneNote.Notebook
        /// Gets the active outline if one exists, If no outline is active, throws ItemNotFound.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveOutline: unit -> OneNote.Outline
        /// Gets the active outline if one exists, otherwise returns null.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveOutlineOrNull: unit -> OneNote.Outline
        /// Gets the active page if one exists. If no page is active, throws ItemNotFound.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActivePage: unit -> OneNote.Page
        /// Gets the active page if one exists. If no page is active, returns null.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActivePageOrNull: unit -> OneNote.Page
        /// Gets the active Paragraph if one exists, If no Paragraph is active, throws ItemNotFound.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveParagraph: unit -> OneNote.Paragraph
        /// Gets the active Paragraph if one exists, otherwise returns null.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveParagraphOrNull: unit -> OneNote.Paragraph
        /// Gets the active section if one exists. If no section is active, throws ItemNotFound.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveSection: unit -> OneNote.Section
        /// Gets the active section if one exists. If no section is active, returns null.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getActiveSectionOrNull: unit -> OneNote.Section
        abstract getWindowSize: unit -> OfficeExtension.ClientResult<ResizeArray<float>>
        abstract insertHtmlAtCurrentPosition: html: string -> unit
        abstract isViewingDeletedNotes: unit -> OfficeExtension.ClientResult<bool>
        /// <summary>Opens the specified page in the application instance.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="page">The page to open.</param>
        abstract navigateToPage: page: OneNote.Page -> unit
        /// <summary>Gets the specified page, and opens it in the application instance.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="url">The client url of the page to open.</param>
        abstract navigateToPageWithClientUrl: url: string -> OneNote.Page
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.ApplicationLoadOptions -> OneNote.Application
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Application
        abstract load: ?option: ApplicationLoadOption -> OneNote.Application
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.ApplicationData

    type [<AllowNullLiteral>] ApplicationLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ApplicationStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Application

    /// Represents ink analysis data for a given set of ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysis =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the parent page object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract page: OneNote.Page
        /// Gets the ID of the InkAnalysis object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.InkAnalysisUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.InkAnalysis -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkAnalysisLoadOptions -> OneNote.InkAnalysis
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysis
        abstract load: ?option: InkAnalysisLoadOption -> OneNote.InkAnalysis
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysis
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysis
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkAnalysis object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisData

    type [<AllowNullLiteral>] InkAnalysisLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents ink analysis data for a given set of ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysis

    /// Represents ink analysis data for an identified paragraph formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisParagraph =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Reference to the parent InkAnalysisPage. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract inkAnalysis: OneNote.InkAnalysis
        /// Gets the ink analysis lines in this ink analysis paragraph. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract lines: OneNote.InkAnalysisLineCollection
        /// Gets the ID of the InkAnalysisParagraph object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.InkAnalysisParagraphUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.InkAnalysisParagraph -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkAnalysisParagraphLoadOptions -> OneNote.InkAnalysisParagraph
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisParagraph
        abstract load: ?option: InkAnalysisParagraphLoadOption -> OneNote.InkAnalysisParagraph
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisParagraph
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisParagraph
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkAnalysisParagraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisParagraphData

    type [<AllowNullLiteral>] InkAnalysisParagraphLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents ink analysis data for an identified paragraph formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisParagraphStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisParagraph

    /// Represents a collection of InkAnalysisParagraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisParagraphCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.InkAnalysisParagraph>
        /// Returns the number of InkAnalysisParagraphs in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a InkAnalysisParagraph object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the InkAnalysisParagraph object, or the index location of the InkAnalysisParagraph object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.InkAnalysisParagraph
        /// <summary>Gets a InkAnalysisParagraph on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.InkAnalysisParagraph
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.InkAnalysisParagraphCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisParagraphCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.InkAnalysisParagraphCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisParagraphCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisParagraphCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.InkAnalysisParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisParagraphCollectionData

    /// Represents a collection of InkAnalysisParagraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisParagraphCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisParagraphCollection

    /// Represents ink analysis data for an identified text line formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisLine =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Reference to the parent InkAnalysisParagraph. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraph: OneNote.InkAnalysisParagraph
        /// Gets the ink analysis words in this ink analysis line. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract words: OneNote.InkAnalysisWordCollection
        /// Gets the ID of the InkAnalysisLine object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.InkAnalysisLineUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.InkAnalysisLine -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkAnalysisLineLoadOptions -> OneNote.InkAnalysisLine
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisLine
        abstract load: ?option: InkAnalysisLineLoadOption -> OneNote.InkAnalysisLine
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisLine
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisLine
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkAnalysisLine object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisLineData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisLineData

    type [<AllowNullLiteral>] InkAnalysisLineLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents ink analysis data for an identified text line formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisLineStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisLine

    /// Represents a collection of InkAnalysisLine objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisLineCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.InkAnalysisLine>
        /// Returns the number of InkAnalysisLines in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a InkAnalysisLine object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the InkAnalysisLine object, or the index location of the InkAnalysisLine object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.InkAnalysisLine
        /// <summary>Gets a InkAnalysisLine on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.InkAnalysisLine
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.InkAnalysisLineCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisLineCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.InkAnalysisLineCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisLineCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisLineCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.InkAnalysisLineCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisLineCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisLineCollectionData

    /// Represents a collection of InkAnalysisLine objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisLineCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisLineCollection

    /// Represents ink analysis data for an identified word formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisWord =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Reference to the parent InkAnalysisLine. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract line: OneNote.InkAnalysisLine
        /// Gets the ID of the InkAnalysisWord object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// The id of the recognized language in this inkAnalysisWord. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract languageId: string
        /// Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract strokePointers: ResizeArray<OneNote.InkStrokePointer>
        /// The words that were recognized in this ink word, in order of likelihood. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract wordAlternates: ResizeArray<string>
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.InkAnalysisWordUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.InkAnalysisWord -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkAnalysisWordLoadOptions -> OneNote.InkAnalysisWord
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisWord
        abstract load: ?option: InkAnalysisWordLoadOption -> OneNote.InkAnalysisWord
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisWord
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisWord
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkAnalysisWord object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisWordData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisWordData

    type [<AllowNullLiteral>] InkAnalysisWordLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents ink analysis data for an identified word formed by ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisWordStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisWord

    /// Represents a collection of InkAnalysisWord objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisWordCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.InkAnalysisWord>
        /// Returns the number of InkAnalysisWords in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the InkAnalysisWord object, or the index location of the InkAnalysisWord object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.InkAnalysisWord
        /// <summary>Gets a InkAnalysisWord on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.InkAnalysisWord
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.InkAnalysisWordCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkAnalysisWordCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.InkAnalysisWordCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkAnalysisWordCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkAnalysisWordCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.InkAnalysisWordCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisWordCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.InkAnalysisWordCollectionData

    /// Represents a collection of InkAnalysisWord objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkAnalysisWordCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkAnalysisWordCollection

    /// Represents a group of ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] FloatingInk =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the strokes of the FloatingInk object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract inkStrokes: OneNote.InkStrokeCollection
        /// Gets the PageContent parent of the FloatingInk object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract pageContent: OneNote.PageContent
        /// Gets the ID of the FloatingInk object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.FloatingInkLoadOptions -> OneNote.FloatingInk
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.FloatingInk
        abstract load: ?option: FloatingInkLoadOption -> OneNote.FloatingInk
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.FloatingInk
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.FloatingInk
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.FloatingInk object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.FloatingInkData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.FloatingInkData

    type [<AllowNullLiteral>] FloatingInkLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a group of ink strokes.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] FloatingInkStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> FloatingInk

    /// Represents a single stroke of ink.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkStroke =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the ID of the InkStroke object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract floatingInk: OneNote.FloatingInk
        /// Gets the ID of the InkStroke object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkStrokeLoadOptions -> OneNote.InkStroke
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkStroke
        abstract load: ?option: InkStrokeLoadOption -> OneNote.InkStroke
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkStroke
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkStroke
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkStroke object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkStrokeData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkStrokeData

    type [<AllowNullLiteral>] InkStrokeLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a single stroke of ink.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkStrokeStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkStroke

    /// Represents a collection of InkStroke objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkStrokeCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.InkStroke>
        /// Returns the number of InkStrokes in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a InkStroke object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the InkStroke object, or the index location of the InkStroke object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.InkStroke
        /// <summary>Gets a InkStroke on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.InkStroke
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.InkStrokeCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkStrokeCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.InkStrokeCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkStrokeCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkStrokeCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.InkStrokeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkStrokeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.InkStrokeCollectionData

    /// Represents a collection of InkStroke objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkStrokeCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkStrokeCollection

    /// A container for the ink in a word in a paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkWord =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// The parent paragraph containing the ink word. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraph: OneNote.Paragraph
        /// Gets the ID of the InkWord object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// The id of the recognized language in this ink word. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract languageId: string
        /// The words that were recognized in this ink word, in order of likelihood. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract wordAlternates: ResizeArray<string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.InkWordLoadOptions -> OneNote.InkWord
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkWord
        abstract load: ?option: InkWordLoadOption -> OneNote.InkWord
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkWord
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkWord
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.InkWord object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkWordData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.InkWordData

    type [<AllowNullLiteral>] InkWordLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// A container for the ink in a word in a paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkWordStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkWord

    /// Represents a collection of InkWord objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkWordCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.InkWord>
        /// Returns the number of InkWords in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a InkWord object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the InkWord object, or the index location of the InkWord object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.InkWord
        /// <summary>Gets a InkWord on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.InkWord
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.InkWordCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.InkWordCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.InkWordCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.InkWordCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.InkWordCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.InkWordCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkWordCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.InkWordCollectionData

    /// Represents a collection of InkWord objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkWordCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InkWordCollection

    /// Represents a OneNote notebook. Notebooks contain section groups and sections.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Notebook =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// The section groups in the notebook. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract sectionGroups: OneNote.SectionGroupCollection
        /// The the sections of the notebook. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract sections: OneNote.SectionCollection
        /// The url of the site that this notebook is located. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract baseUrl: string
        /// The client url of the notebook. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clientUrl: string
        /// Gets the ID of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
        /// 
        /// [Api set: OneNoteApi 1.2]
        abstract isVirtual: bool
        /// Gets the name of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract name: string
        /// <summary>Adds a new section to the end of the notebook.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the new section.</param>
        abstract addSection: name: string -> OneNote.Section
        /// <summary>Adds a new section group to the end of the notebook.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the new section.</param>
        abstract addSectionGroup: name: string -> OneNote.SectionGroup
        /// Gets the REST API ID.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getRestApiId: unit -> OfficeExtension.ClientResult<string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.NotebookLoadOptions -> OneNote.Notebook
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Notebook
        abstract load: ?option: NotebookLoadOption -> OneNote.Notebook
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Notebook
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Notebook
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Notebook object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NotebookData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.NotebookData

    type [<AllowNullLiteral>] NotebookLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a OneNote notebook. Notebooks contain section groups and sections.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] NotebookStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Notebook

    /// Represents a collection of notebooks.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] NotebookCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.Notebook>
        /// Returns the number of notebooks in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets the collection of notebooks with the specified name that are open in the application instance.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the notebook.</param>
        abstract getByName: name: string -> OneNote.NotebookCollection
        /// <summary>Gets a notebook by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the notebook, or the index location of the notebook in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.Notebook
        /// <summary>Gets a notebook on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.Notebook
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.NotebookCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.NotebookCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.NotebookCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.NotebookCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.NotebookCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.NotebookCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NotebookCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.NotebookCollectionData

    /// Represents a collection of notebooks.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] NotebookCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> NotebookCollection

    /// Represents a OneNote section group. Section groups can contain sections and other section groups.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionGroup =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the notebook that contains the section group. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract notebook: OneNote.Notebook
        /// Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentSectionGroup: OneNote.SectionGroup
        /// Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentSectionGroupOrNull: OneNote.SectionGroup
        /// The collection of section groups in the section group. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract sectionGroups: OneNote.SectionGroupCollection
        /// The collection of sections in the section group. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract sections: OneNote.SectionCollection
        /// The client url of the section group. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clientUrl: string
        /// Gets the ID of the section group. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the name of the section group. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract name: string
        /// <summary>Adds a new section to the end of the section group.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="title">The name of the new section.</param>
        abstract addSection: title: string -> OneNote.Section
        /// <summary>Adds a new section group to the end of this sectionGroup.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the new section.</param>
        abstract addSectionGroup: name: string -> OneNote.SectionGroup
        /// Gets the REST API ID.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getRestApiId: unit -> OfficeExtension.ClientResult<string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.SectionGroupLoadOptions -> OneNote.SectionGroup
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.SectionGroup
        abstract load: ?option: SectionGroupLoadOption -> OneNote.SectionGroup
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.SectionGroup
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.SectionGroup
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.SectionGroup object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionGroupData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.SectionGroupData

    type [<AllowNullLiteral>] SectionGroupLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a OneNote section group. Section groups can contain sections and other section groups.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionGroupStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SectionGroup

    /// Represents a collection of section groups.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionGroupCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.SectionGroup>
        /// Returns the number of section groups in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets the collection of section groups with the specified name.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the section group.</param>
        abstract getByName: name: string -> OneNote.SectionGroupCollection
        /// <summary>Gets a section group by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the section group, or the index location of the section group in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.SectionGroup
        /// <summary>Gets a section group on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.SectionGroup
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.SectionGroupCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.SectionGroupCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.SectionGroupCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.SectionGroupCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.SectionGroupCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.SectionGroupCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionGroupCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.SectionGroupCollectionData

    /// Represents a collection of section groups.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionGroupCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SectionGroupCollection

    /// Represents a OneNote section. Sections can contain pages.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Section =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the notebook that contains the section. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract notebook: OneNote.Notebook
        /// The collection of pages in the section. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract pages: OneNote.PageCollection
        /// Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentSectionGroup: OneNote.SectionGroup
        /// Gets the section group that contains the section. Returns null if the section is a direct child of the notebook. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentSectionGroupOrNull: OneNote.SectionGroup
        /// The client url of the section. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clientUrl: string
        /// Gets the ID of the section. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// True if this section is encrypted with a password. Read only
        /// 
        /// [Api set: OneNoteApi 1.2]
        abstract isEncrypted: bool
        /// True if this section is locked. Read only
        /// 
        /// [Api set: OneNoteApi 1.2]
        abstract isLocked: bool
        /// Gets the name of the section. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract name: string
        /// The web url of the page. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract webUrl: string
        /// <summary>Adds a new page to the end of the section.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="title">The title of the new page.</param>
        abstract addPage: title: string -> OneNote.Page
        /// <summary>Copies this section to specified notebook.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="destinationNotebook">The notebook to copy this section to.</param>
        abstract copyToNotebook: destinationNotebook: OneNote.Notebook -> OneNote.Section
        /// <summary>Copies this section to specified section group.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="destinationSectionGroup">The section group to copy this section to.</param>
        abstract copyToSectionGroup: destinationSectionGroup: OneNote.SectionGroup -> OneNote.Section
        /// Gets the REST API ID.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getRestApiId: unit -> OfficeExtension.ClientResult<string>
        /// <summary>Inserts a new section before or after the current section.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="location">The location of the new section relative to the current section.</param>
        /// <param name="title">The name of the new section.</param>
        abstract insertSectionAsSibling: location: OneNote.InsertLocation * title: string -> OneNote.Section
        /// <summary>Inserts a new section before or after the current section.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="location">The location of the new section relative to the current section.</param>
        /// <param name="title">The name of the new section.</param>
        abstract insertSectionAsSibling: location: SectionInsertSectionAsSiblingLocation * title: string -> OneNote.Section
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.SectionLoadOptions -> OneNote.Section
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Section
        abstract load: ?option: SectionLoadOption -> OneNote.Section
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Section
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Section
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Section object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.SectionData

    type [<StringEnum>] [<RequireQualifiedAccess>] SectionInsertSectionAsSiblingLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<AllowNullLiteral>] SectionLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a OneNote section. Sections can contain pages.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Section

    /// Represents a collection of sections.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.Section>
        /// Returns the number of sections in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets the collection of sections with the specified name.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="name">The name of the section.</param>
        abstract getByName: name: string -> OneNote.SectionCollection
        /// <summary>Gets a section by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the section, or the index location of the section in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.Section
        /// <summary>Gets a section on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.Section
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.SectionCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.SectionCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.SectionCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.SectionCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.SectionCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.SectionCollectionData

    /// Represents a collection of sections.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] SectionCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SectionCollection

    /// Represents a OneNote page.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Page =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// The collection of PageContent objects on the page. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract contents: OneNote.PageContentCollection
        /// Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract inkAnalysisOrNull: OneNote.InkAnalysis
        /// Gets the section that contains the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentSection: OneNote.Section
        /// Gets the ClassNotebookPageSource to the page.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract classNotebookPageSource: string
        /// The client url of the page. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clientUrl: string
        /// Gets the ID of the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets or sets the indentation level of the page.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract pageLevel: float with get, set
        /// Gets or sets the title of the page.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract title: string with get, set
        /// The web url of the page. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract webUrl: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.PageUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.Page -> unit
        /// <summary>Adds an Outline to the page at the specified position.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="left">The left position of the top, left corner of the Outline.</param>
        /// <param name="top">The top position of the top, left corner of the Outline.</param>
        /// <param name="html">An HTML string that describes the visual presentation of the Outline. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.</param>
        abstract addOutline: left: float * top: float * html: string -> OneNote.Outline
        /// Return a json string with node id and content in html format.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract analyzePage: unit -> OfficeExtension.ClientResult<string>
        /// <summary>Inserts a new page with translated content.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="translatedContent">Translated content of the page</param>
        abstract applyTranslation: translatedContent: string -> unit
        /// <summary>Copies this page to specified section.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="destinationSection">The section to copy this page to.</param>
        abstract copyToSection: destinationSection: OneNote.Section -> OneNote.Page
        /// Copies this page to specified section and sets ClassNotebookPageSource.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract copyToSectionAndSetClassNotebookPageSource: destinationSection: OneNote.Section -> OneNote.Page
        /// Gets the REST API ID.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getRestApiId: unit -> OfficeExtension.ClientResult<string>
        /// Does the page has content title.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract hasTitleContent: unit -> OfficeExtension.ClientResult<bool>
        /// <summary>Inserts a new page before or after the current page.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="location">The location of the new page relative to the current page.</param>
        /// <param name="title">The title of the new page.</param>
        abstract insertPageAsSibling: location: OneNote.InsertLocation * title: string -> OneNote.Page
        /// <summary>Inserts a new page before or after the current page.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="location">The location of the new page relative to the current page.</param>
        /// <param name="title">The title of the new page.</param>
        abstract insertPageAsSibling: location: PageInsertPageAsSiblingLocation * title: string -> OneNote.Page
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.PageLoadOptions -> OneNote.Page
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Page
        abstract load: ?option: PageLoadOption -> OneNote.Page
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Page
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Page
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Page object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.PageData

    type [<StringEnum>] [<RequireQualifiedAccess>] PageInsertPageAsSiblingLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<AllowNullLiteral>] PageLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a OneNote page.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Page

    /// Represents a collection of pages.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.Page>
        /// Returns the number of pages in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets the collection of pages with the specified title.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="title">The title of the page.</param>
        abstract getByTitle: title: string -> OneNote.PageCollection
        /// <summary>Gets a page by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the page, or the index location of the page in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.Page
        /// <summary>Gets a page on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.Page
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.PageCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.PageCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.PageCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.PageCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.PageCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.PageCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.PageCollectionData

    /// Represents a collection of pages.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> PageCollection

    /// Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageContent =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract image: OneNote.Image
        /// Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ink: OneNote.FloatingInk
        /// Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract outline: OneNote.Outline
        /// Gets the page that contains the PageContent object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentPage: OneNote.Page
        /// Gets the ID of the PageContent object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets or sets the left (X-axis) position of the PageContent object.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract left: float with get, set
        /// Gets or sets the top (Y-axis) position of the PageContent object.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract top: float with get, set
        /// Gets the type of the PageContent object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ``type``: U2<OneNote.PageContentType, string>
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.PageContentUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.PageContent -> unit
        /// Deletes the PageContent object.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract delete: unit -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.PageContentLoadOptions -> OneNote.PageContent
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.PageContent
        abstract load: ?option: PageContentLoadOption -> OneNote.PageContent
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.PageContent
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.PageContent
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.PageContent object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageContentData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.PageContentData

    type [<AllowNullLiteral>] PageContentLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageContentStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> PageContent

    /// Represents the contents of a page, as a collection of PageContent objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageContentCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.PageContent>
        /// Returns the number of page contents in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a PageContent object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the PageContent object, or the index location of the PageContent object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.PageContent
        /// <summary>Gets a page content on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.PageContent
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.PageContentCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.PageContentCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.PageContentCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.PageContentCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.PageContentCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.PageContentCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageContentCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.PageContentCollectionData

    /// Represents the contents of a page, as a collection of PageContent objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] PageContentCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> PageContentCollection

    /// Represents a container for Paragraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Outline =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract pageContent: OneNote.PageContent
        /// Gets the collection of Paragraph objects in the Outline. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraphs: OneNote.ParagraphCollection
        /// Gets the ID of the Outline object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// <summary>Adds the specified HTML to the bottom of the Outline.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="html">The HTML string to append. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.</param>
        abstract appendHtml: html: string -> unit
        /// <summary>Adds the specified image to the bottom of the Outline.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="base64EncodedImage">HTML string to append.</param>
        /// <param name="width">Optional. Width in the unit of Points. The default value is null and image width will be respected.</param>
        /// <param name="height">Optional. Height in the unit of Points. The default value is null and image height will be respected.</param>
        abstract appendImage: base64EncodedImage: string * width: float * height: float -> OneNote.Image
        /// <summary>Adds the specified text to the bottom of the Outline.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="paragraphText">HTML string to append.</param>
        abstract appendRichText: paragraphText: string -> OneNote.RichText
        /// <summary>Adds a table with the specified number of rows and columns to the bottom of the outline.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract appendTable: rowCount: float * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> OneNote.Table
        /// Check if the outline is title outline.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract isTitle: unit -> OfficeExtension.ClientResult<bool>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.OutlineLoadOptions -> OneNote.Outline
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Outline
        abstract load: ?option: OutlineLoadOption -> OneNote.Outline
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Outline
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Outline
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Outline object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.OutlineData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.OutlineData

    type [<AllowNullLiteral>] OutlineLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a container for Paragraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] OutlineStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Outline

    /// A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Paragraph =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract image: OneNote.Image
        /// Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract inkWords: OneNote.InkWordCollection
        /// Gets the Outline object that contains the Paragraph. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract outline: OneNote.Outline
        /// The collection of paragraphs under this paragraph. Read only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraphs: OneNote.ParagraphCollection
        /// Gets the parent paragraph object. Throws if a parent paragraph does not exist. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentParagraph: OneNote.Paragraph
        /// Gets the parent paragraph object. Returns null if a parent paragraph does not exist. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentParagraphOrNull: OneNote.Paragraph
        /// Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentTableCell: OneNote.TableCell
        /// Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentTableCellOrNull: OneNote.TableCell
        /// Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract richText: OneNote.RichText
        /// Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract table: OneNote.Table
        /// Gets the ID of the Paragraph object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the type of the Paragraph object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ``type``: U2<OneNote.ParagraphType, string>
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.ParagraphUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.Paragraph -> unit
        /// <summary>Add NoteTag to the paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="type">The type of the NoteTag.</param>
        /// <param name="status">The status of the NoteTag.</param>
        abstract addNoteTag: ``type``: OneNote.NoteTagType * status: OneNote.NoteTagStatus -> OneNote.NoteTag
        /// <summary>Add NoteTag to the paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="type">The type of the NoteTag.</param>
        /// <param name="status">The status of the NoteTag.</param>
        abstract addNoteTag: ``type``: ParagraphAddNoteTagType * status: ParagraphAddNoteTagStatus -> OneNote.NoteTag
        /// Deletes the paragraph
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract delete: unit -> unit
        /// Get list information of paragraph
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getParagraphInfo: unit -> OfficeExtension.ClientResult<OneNote.ParagraphInfo>
        /// <summary>Inserts the specified HTML content
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of new contents relative to the current Paragraph.</param>
        /// <param name="html">An HTML string that describes the visual presentation of the content. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.</param>
        abstract insertHtmlAsSibling: insertLocation: OneNote.InsertLocation * html: string -> unit
        /// <summary>Inserts the specified HTML content
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of new contents relative to the current Paragraph.</param>
        /// <param name="html">An HTML string that describes the visual presentation of the content. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.</param>
        abstract insertHtmlAsSibling: insertLocation: ParagraphInsertHtmlAsSiblingInsertLocation * html: string -> unit
        /// <summary>Inserts the image at the specified insert location..
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="base64EncodedImage">HTML string to append.</param>
        /// <param name="width">Optional. Width in the unit of Points. The default value is null and image width will be respected.</param>
        /// <param name="height">Optional. Height in the unit of Points. The default value is null and image height will be respected.</param>
        abstract insertImageAsSibling: insertLocation: OneNote.InsertLocation * base64EncodedImage: string * width: float * height: float -> OneNote.Image
        /// <summary>Inserts the image at the specified insert location..
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="base64EncodedImage">HTML string to append.</param>
        /// <param name="width">Optional. Width in the unit of Points. The default value is null and image width will be respected.</param>
        /// <param name="height">Optional. Height in the unit of Points. The default value is null and image height will be respected.</param>
        abstract insertImageAsSibling: insertLocation: ParagraphInsertImageAsSiblingInsertLocation * base64EncodedImage: string * width: float * height: float -> OneNote.Image
        /// <summary>Inserts the paragraph text at the specifiec insert location.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="paragraphText">HTML string to append.</param>
        abstract insertRichTextAsSibling: insertLocation: OneNote.InsertLocation * paragraphText: string -> OneNote.RichText
        /// <summary>Inserts the paragraph text at the specifiec insert location.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="paragraphText">HTML string to append.</param>
        abstract insertRichTextAsSibling: insertLocation: ParagraphInsertRichTextAsSiblingInsertLocation * paragraphText: string -> OneNote.RichText
        /// <summary>Adds a table with the specified number of rows and columns before or after the current paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="rowCount">The number of rows in the table.</param>
        /// <param name="columnCount">The number of columns in the table.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTableAsSibling: insertLocation: OneNote.InsertLocation * rowCount: float * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> OneNote.Table
        /// <summary>Adds a table with the specified number of rows and columns before or after the current paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">The location of the table relative to the current Paragraph.</param>
        /// <param name="rowCount">The number of rows in the table.</param>
        /// <param name="columnCount">The number of columns in the table.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTableAsSibling: insertLocation: ParagraphInsertTableAsSiblingInsertLocation * rowCount: float * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> OneNote.Table
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.ParagraphLoadOptions -> OneNote.Paragraph
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Paragraph
        abstract load: ?option: ParagraphLoadOption -> OneNote.Paragraph
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Paragraph
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Paragraph
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Paragraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.ParagraphData

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphAddNoteTagType =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "ToDo">] ToDo
        | [<CompiledName "Important">] Important
        | [<CompiledName "Question">] Question
        | [<CompiledName "Contact">] Contact
        | [<CompiledName "Address">] Address
        | [<CompiledName "PhoneNumber">] PhoneNumber
        | [<CompiledName "Website">] Website
        | [<CompiledName "Idea">] Idea
        | [<CompiledName "Critical">] Critical
        | [<CompiledName "ToDoPriority1">] ToDoPriority1
        | [<CompiledName "ToDoPriority2">] ToDoPriority2

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphAddNoteTagStatus =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "Normal">] Normal
        | [<CompiledName "Completed">] Completed
        | [<CompiledName "Disabled">] Disabled
        | [<CompiledName "OutlookTask">] OutlookTask
        | [<CompiledName "TaskNotSyncedYet">] TaskNotSyncedYet
        | [<CompiledName "TaskRemoved">] TaskRemoved

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertHtmlAsSiblingInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertImageAsSiblingInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertRichTextAsSiblingInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertTableAsSiblingInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<AllowNullLiteral>] ParagraphLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ParagraphStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Paragraph

    /// Represents a collection of Paragraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ParagraphCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.Paragraph>
        /// Returns the number of paragraphs in the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a Paragraph object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">The ID of the Paragraph object, or the index location of the Paragraph object in the collection.</param>
        abstract getItem: index: U2<float, string> -> OneNote.Paragraph
        /// <summary>Gets a paragraph on its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.Paragraph
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.ParagraphCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.ParagraphCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.ParagraphCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.ParagraphCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.ParagraphCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.ParagraphCollectionData

    /// Represents a collection of Paragraph objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ParagraphCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ParagraphCollection

    /// A container for the NoteTag in a paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] NoteTag =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the Id of the NoteTag object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the status of the NoteTag object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract status: U2<OneNote.NoteTagStatus, string>
        /// Gets the type of the NoteTag object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ``type``: U2<OneNote.NoteTagType, string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.NoteTagLoadOptions -> OneNote.NoteTag
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.NoteTag
        abstract load: ?option: NoteTagLoadOption -> OneNote.NoteTag
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.NoteTag
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.NoteTag
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.NoteTag object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NoteTagData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.NoteTagData

    type [<AllowNullLiteral>] NoteTagLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// A container for the NoteTag in a paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] NoteTagStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> NoteTag

    /// Represents a RichText object in a Paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] RichText =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the Paragraph object that contains the RichText object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraph: OneNote.Paragraph
        /// Gets the ID of the RichText object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// The language id of the text. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract languageId: string
        /// Gets the text content of the RichText object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract text: string
        /// Get the HTML of the rich text
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getHtml: unit -> OfficeExtension.ClientResult<string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.RichTextLoadOptions -> OneNote.RichText
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.RichText
        abstract load: ?option: RichTextLoadOption -> OneNote.RichText
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.RichText
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.RichText
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.RichText object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.RichTextData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.RichTextData

    type [<AllowNullLiteral>] RichTextLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a RichText object in a Paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] RichTextStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> RichText

    /// Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Image =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract pageContent: OneNote.PageContent
        /// Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraph: OneNote.Paragraph
        /// Gets or sets the description of the Image.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract description: string with get, set
        /// Gets or sets the height of the Image layout.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract height: float with get, set
        /// Gets or sets the hyperlink of the Image.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract hyperlink: string with get, set
        /// Gets the ID of the Image object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ocrData: OneNote.ImageOcrData
        /// Gets or sets the width of the Image layout.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract width: float with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.ImageUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.Image -> unit
        /// Gets the base64-encoded binary representation of the Image.
        ///   Example: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIA...
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract getBase64Image: unit -> OfficeExtension.ClientResult<string>
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.ImageLoadOptions -> OneNote.Image
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Image
        abstract load: ?option: ImageLoadOption -> OneNote.Image
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Image
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Image
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Image object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ImageData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.ImageData

    type [<AllowNullLiteral>] ImageLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ImageStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Image

    /// Represents a table in a OneNote page.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] Table =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the Paragraph object that contains the Table object. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraph: OneNote.Paragraph
        /// Gets all of the table rows. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract rows: OneNote.TableRowCollection
        /// Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract borderVisible: bool with get, set
        /// Gets the number of columns in the table.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract columnCount: float
        /// Gets the ID of the table. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the number of rows in the table.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract rowCount: float
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.Table -> unit
        /// <summary>Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="values">Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.</param>
        abstract appendColumn: ?values: ResizeArray<string> -> unit
        /// <summary>Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="values">Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.</param>
        abstract appendRow: ?values: ResizeArray<string> -> OneNote.TableRow
        /// Clears the contents of the table.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clear: unit -> unit
        /// <summary>Gets the table cell at a specified row and column.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="rowIndex">The index of the row.</param>
        /// <param name="cellIndex">The index of the cell in the row.</param>
        abstract getCell: rowIndex: float * cellIndex: float -> OneNote.TableCell
        /// <summary>Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index where the column will be inserted in the table.</param>
        /// <param name="values">Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.</param>
        abstract insertColumn: index: float * ?values: ResizeArray<string> -> unit
        /// <summary>Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index where the row will be inserted in the table.</param>
        /// <param name="values">Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.</param>
        abstract insertRow: index: float * ?values: ResizeArray<string> -> OneNote.TableRow
        /// Sets the shading color of all cells in the table.
        ///   The color code to set the cells to.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract setShadingColor: colorCode: string -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.TableLoadOptions -> OneNote.Table
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.Table
        abstract load: ?option: TableLoadOption -> OneNote.Table
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.Table
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.Table
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.Table object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.TableData

    type [<AllowNullLiteral>] TableLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a table in a OneNote page.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Table

    /// Represents a row in a table.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableRow =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the cells in the row. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract cells: OneNote.TableCellCollection
        /// Gets the parent table. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentTable: OneNote.Table
        /// Gets the number of cells in the row. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract cellCount: float
        /// Gets the ID of the row. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the index of the row in its parent table. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract rowIndex: float
        /// Clears the contents of the row.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clear: unit -> unit
        /// <summary>Inserts a row before or after the current row.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">Where the new rows should be inserted relative to the current row.</param>
        /// <param name="values">Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.</param>
        abstract insertRowAsSibling: insertLocation: OneNote.InsertLocation * ?values: ResizeArray<string> -> OneNote.TableRow
        /// <summary>Inserts a row before or after the current row.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="insertLocation">Where the new rows should be inserted relative to the current row.</param>
        /// <param name="values">Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.</param>
        abstract insertRowAsSibling: insertLocation: TableRowInsertRowAsSiblingInsertLocation * ?values: ResizeArray<string> -> OneNote.TableRow
        /// Sets the shading color of all cells in the row.
        ///   The color code to set the cells to.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract setShadingColor: colorCode: string -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.TableRowLoadOptions -> OneNote.TableRow
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.TableRow
        abstract load: ?option: TableRowLoadOption -> OneNote.TableRow
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.TableRow
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.TableRow
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.TableRow object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.TableRowData

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowInsertRowAsSiblingInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<AllowNullLiteral>] TableRowLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a row in a table.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableRowStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableRow

    /// Contains a collection of TableRow objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableRowCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.TableRow>
        /// Returns the number of table rows in this collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a table row object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">A number that identifies the index location of a table row object.</param>
        abstract getItem: index: U2<float, string> -> OneNote.TableRow
        /// <summary>Gets a table row at its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.TableRow
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.TableRowCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.TableRowCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.TableRowCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.TableRowCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.TableRowCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.TableRowCollectionData

    /// Contains a collection of TableRow objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableRowCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableRowCollection

    /// Represents a cell in a OneNote table.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableCell =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of Paragraph objects in the TableCell. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract paragraphs: OneNote.ParagraphCollection
        /// Gets the parent row of the cell. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract parentRow: OneNote.TableRow
        /// Gets the index of the cell in its row. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract cellIndex: float
        /// Gets the ID of the cell. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract id: string
        /// Gets the index of the cell's row in the table. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract rowIndex: float
        /// Gets and sets the shading color of the cell
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract shadingColor: string with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableCellUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: OneNote.TableCell -> unit
        /// <summary>Adds the specified HTML to the bottom of the TableCell.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="html">The HTML string to append. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.</param>
        abstract appendHtml: html: string -> unit
        /// <summary>Adds the specified image to table cell.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="base64EncodedImage">HTML string to append.</param>
        /// <param name="width">Optional. Width in the unit of Points. The default value is null and image width will be respected.</param>
        /// <param name="height">Optional. Height in the unit of Points. The default value is null and image height will be respected.</param>
        abstract appendImage: base64EncodedImage: string * width: float * height: float -> OneNote.Image
        /// <summary>Adds the specified text to table cell.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="paragraphText">HTML string to append.</param>
        abstract appendRichText: paragraphText: string -> OneNote.RichText
        /// <summary>Adds a table with the specified number of rows and columns to table cell.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract appendTable: rowCount: float * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> OneNote.Table
        /// Clears the contents of the cell.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract clear: unit -> unit
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: OneNote.Interfaces.TableCellLoadOptions -> OneNote.TableCell
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.TableCell
        abstract load: ?option: TableCellLoadOption -> OneNote.TableCell
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.TableCell
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.TableCell
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original OneNote.TableCell object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> OneNote.Interfaces.TableCellData

    type [<AllowNullLiteral>] TableCellLoadOption =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a cell in a OneNote table.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableCellStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableCell

    /// Contains a collection of TableCell objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableCellCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<OneNote.TableCell>
        /// Returns the number of tablecells in this collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract count: float
        /// <summary>Gets a table cell object by ID or by its index in the collection. Read-only.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">A number that identifies the index location of a table cell object.</param>
        abstract getItem: index: U2<float, string> -> OneNote.TableCell
        /// <summary>Gets a tablecell at its position in the collection.
        /// 
        /// [Api set: OneNoteApi 1.1]</summary>
        /// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
        abstract getItemAt: index: float -> OneNote.TableCell
        /// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
        abstract load: ?option: obj -> OneNote.TableCellCollection
        abstract load: ?option: U2<string, ResizeArray<string>> -> OneNote.TableCellCollection
        abstract load: ?option: OfficeExtension.LoadOption -> OneNote.TableCellCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> OneNote.TableCellCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
        abstract untrack: unit -> OneNote.TableCellCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `OneNote.TableCellCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableCellCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> OneNote.Interfaces.TableCellCollectionData

    /// Contains a collection of TableCell objects.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] TableCellCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableCellCollection

    /// Represents data obtained by OCR (optical character recognition) of an image.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ImageOcrData =
        /// Represents the OCR language, with values such as EN-US
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ocrLanguageId: string with get, set
        /// Represents the text obtained by OCR of the image
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract ocrText: string with get, set

    /// Weak reference to an ink stroke object and its content parent.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] InkStrokePointer =
        /// Represents the id of the page content object corresponding to this stroke
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract contentId: string with get, set
        /// Represents the id of the ink stroke
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract inkStrokeId: string with get, set

    /// List information for paragraph.
    /// 
    /// [Api set: OneNoteApi 1.1]
    type [<AllowNullLiteral>] ParagraphInfo =
        /// //
        ///   Bullet list type of paragraph
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract bulletType: string with get, set
        /// //
        ///   Index of paragraph in list
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract index: float with get, set
        /// //
        ///   Type of list in paragraph
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract listType: U2<OneNote.ListType, string> with get, set
        /// //
        ///   number list type of paragraph
        /// 
        /// [Api set: OneNoteApi 1.1]
        abstract numberType: U2<OneNote.NumberType, string> with get, set

    type [<StringEnum>] [<RequireQualifiedAccess>] InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After

    type [<StringEnum>] [<RequireQualifiedAccess>] PageContentType =
        | [<CompiledName "Outline">] Outline
        | [<CompiledName "Image">] Image
        | [<CompiledName "Ink">] Ink
        | [<CompiledName "Other">] Other

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphType =
        | [<CompiledName "RichText">] RichText
        | [<CompiledName "Image">] Image
        | [<CompiledName "Table">] Table
        | [<CompiledName "Ink">] Ink
        | [<CompiledName "Other">] Other

    type [<StringEnum>] [<RequireQualifiedAccess>] NoteTagType =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "ToDo">] ToDo
        | [<CompiledName "Important">] Important
        | [<CompiledName "Question">] Question
        | [<CompiledName "Contact">] Contact
        | [<CompiledName "Address">] Address
        | [<CompiledName "PhoneNumber">] PhoneNumber
        | [<CompiledName "Website">] Website
        | [<CompiledName "Idea">] Idea
        | [<CompiledName "Critical">] Critical
        | [<CompiledName "ToDoPriority1">] ToDoPriority1
        | [<CompiledName "ToDoPriority2">] ToDoPriority2

    type [<StringEnum>] [<RequireQualifiedAccess>] NoteTagStatus =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "Normal">] Normal
        | [<CompiledName "Completed">] Completed
        | [<CompiledName "Disabled">] Disabled
        | [<CompiledName "OutlookTask">] OutlookTask
        | [<CompiledName "TaskNotSyncedYet">] TaskNotSyncedYet
        | [<CompiledName "TaskRemoved">] TaskRemoved

    type [<StringEnum>] [<RequireQualifiedAccess>] ListType =
        | [<CompiledName "None">] None
        | [<CompiledName "Number">] Number
        | [<CompiledName "Bullet">] Bullet

    type [<StringEnum>] [<RequireQualifiedAccess>] NumberType =
        | [<CompiledName "None">] None
        | [<CompiledName "Arabic">] Arabic
        | [<CompiledName "UCRoman">] Ucroman
        | [<CompiledName "LCRoman">] Lcroman
        | [<CompiledName "UCLetter">] Ucletter
        | [<CompiledName "LCLetter">] Lcletter
        | [<CompiledName "Ordinal">] Ordinal
        | [<CompiledName "Cardtext">] Cardtext
        | [<CompiledName "Ordtext">] Ordtext
        | [<CompiledName "Hex">] Hex
        | [<CompiledName "ChiManSty">] ChiManSty
        | [<CompiledName "DbNum1">] DbNum1
        | [<CompiledName "DbNum2">] DbNum2
        | [<CompiledName "Aiueo">] Aiueo
        | [<CompiledName "Iroha">] Iroha
        | [<CompiledName "DbChar">] DbChar
        | [<CompiledName "SbChar">] SbChar
        | [<CompiledName "DbNum3">] DbNum3
        | [<CompiledName "DbNum4">] DbNum4
        | [<CompiledName "Circlenum">] Circlenum
        | [<CompiledName "DArabic">] Darabic
        | [<CompiledName "DAiueo">] Daiueo
        | [<CompiledName "DIroha">] Diroha
        | [<CompiledName "ArabicLZ">] ArabicLZ
        | [<CompiledName "Bullet">] Bullet
        | [<CompiledName "Ganada">] Ganada
        | [<CompiledName "Chosung">] Chosung
        | [<CompiledName "GB1">] Gb1
        | [<CompiledName "GB2">] Gb2
        | [<CompiledName "GB3">] Gb3
        | [<CompiledName "GB4">] Gb4
        | [<CompiledName "Zodiac1">] Zodiac1
        | [<CompiledName "Zodiac2">] Zodiac2
        | [<CompiledName "Zodiac3">] Zodiac3
        | [<CompiledName "TpeDbNum1">] TpeDbNum1
        | [<CompiledName "TpeDbNum2">] TpeDbNum2
        | [<CompiledName "TpeDbNum3">] TpeDbNum3
        | [<CompiledName "TpeDbNum4">] TpeDbNum4
        | [<CompiledName "ChnDbNum1">] ChnDbNum1
        | [<CompiledName "ChnDbNum2">] ChnDbNum2
        | [<CompiledName "ChnDbNum3">] ChnDbNum3
        | [<CompiledName "ChnDbNum4">] ChnDbNum4
        | [<CompiledName "KorDbNum1">] KorDbNum1
        | [<CompiledName "KorDbNum2">] KorDbNum2
        | [<CompiledName "KorDbNum3">] KorDbNum3
        | [<CompiledName "KorDbNum4">] KorDbNum4
        | [<CompiledName "Hebrew1">] Hebrew1
        | [<CompiledName "Arabic1">] Arabic1
        | [<CompiledName "Hebrew2">] Hebrew2
        | [<CompiledName "Arabic2">] Arabic2
        | [<CompiledName "Hindi1">] Hindi1
        | [<CompiledName "Hindi2">] Hindi2
        | [<CompiledName "Hindi3">] Hindi3
        | [<CompiledName "Thai1">] Thai1
        | [<CompiledName "Thai2">] Thai2
        | [<CompiledName "NumInDash">] NumInDash
        | [<CompiledName "LCRus">] Lcrus
        | [<CompiledName "UCRus">] Ucrus
        | [<CompiledName "LCGreek">] Lcgreek
        | [<CompiledName "UCGreek">] Ucgreek
        | [<CompiledName "Lim">] Lim
        | [<CompiledName "Custom">] Custom

    type [<StringEnum>] [<RequireQualifiedAccess>] ErrorCodes =
        | [<CompiledName "GeneralException">] GeneralException

    module Interfaces =

        /// Provides ways to load properties of only a subset of members of a collection.
        type [<AllowNullLiteral>] CollectionLoadOptions =
            /// Specify the number of items in the queried collection to be included in the result.
            abstract ``$top``: float option with get, set
            /// Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            abstract ``$skip``: float option with get, set

        /// An interface for updating data on the InkAnalysis object, for use in "inkAnalysis.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisUpdateData =
            /// Gets the parent page object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract page: OneNote.Interfaces.PageUpdateData option with get, set

        /// An interface for updating data on the InkAnalysisParagraph object, for use in "inkAnalysisParagraph.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisParagraphUpdateData =
            /// Reference to the parent InkAnalysisPage.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysis: OneNote.Interfaces.InkAnalysisUpdateData option with get, set

        /// An interface for updating data on the InkAnalysisParagraphCollection object, for use in "inkAnalysisParagraphCollection.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisParagraphCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisParagraphData> option with get, set

        /// An interface for updating data on the InkAnalysisLine object, for use in "inkAnalysisLine.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisLineUpdateData =
            /// Reference to the parent InkAnalysisParagraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.InkAnalysisParagraphUpdateData option with get, set

        /// An interface for updating data on the InkAnalysisLineCollection object, for use in "inkAnalysisLineCollection.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisLineCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisLineData> option with get, set

        /// An interface for updating data on the InkAnalysisWord object, for use in "inkAnalysisWord.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisWordUpdateData =
            /// Reference to the parent InkAnalysisLine.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract line: OneNote.Interfaces.InkAnalysisLineUpdateData option with get, set

        /// An interface for updating data on the InkAnalysisWordCollection object, for use in "inkAnalysisWordCollection.set({ ... })".
        type [<AllowNullLiteral>] InkAnalysisWordCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisWordData> option with get, set

        /// An interface for updating data on the InkStrokeCollection object, for use in "inkStrokeCollection.set({ ... })".
        type [<AllowNullLiteral>] InkStrokeCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.InkStrokeData> option with get, set

        /// An interface for updating data on the InkWordCollection object, for use in "inkWordCollection.set({ ... })".
        type [<AllowNullLiteral>] InkWordCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.InkWordData> option with get, set

        /// An interface for updating data on the NotebookCollection object, for use in "notebookCollection.set({ ... })".
        type [<AllowNullLiteral>] NotebookCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.NotebookData> option with get, set

        /// An interface for updating data on the SectionGroupCollection object, for use in "sectionGroupCollection.set({ ... })".
        type [<AllowNullLiteral>] SectionGroupCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.SectionGroupData> option with get, set

        /// An interface for updating data on the SectionCollection object, for use in "sectionCollection.set({ ... })".
        type [<AllowNullLiteral>] SectionCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.SectionData> option with get, set

        /// An interface for updating data on the Page object, for use in "page.set({ ... })".
        type [<AllowNullLiteral>] PageUpdateData =
            /// Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysisOrNull: OneNote.Interfaces.InkAnalysisUpdateData option with get, set
            /// Gets or sets the indentation level of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageLevel: float option with get, set
            /// Gets or sets the title of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract title: string option with get, set

        /// An interface for updating data on the PageCollection object, for use in "pageCollection.set({ ... })".
        type [<AllowNullLiteral>] PageCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.PageData> option with get, set

        /// An interface for updating data on the PageContent object, for use in "pageContent.set({ ... })".
        type [<AllowNullLiteral>] PageContentUpdateData =
            /// Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageUpdateData option with get, set
            /// Gets or sets the left (X-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract left: float option with get, set
            /// Gets or sets the top (Y-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract top: float option with get, set

        /// An interface for updating data on the PageContentCollection object, for use in "pageContentCollection.set({ ... })".
        type [<AllowNullLiteral>] PageContentCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.PageContentData> option with get, set

        /// An interface for updating data on the Paragraph object, for use in "paragraph.set({ ... })".
        type [<AllowNullLiteral>] ParagraphUpdateData =
            /// Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageUpdateData option with get, set
            /// Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract table: OneNote.Interfaces.TableUpdateData option with get, set

        /// An interface for updating data on the ParagraphCollection object, for use in "paragraphCollection.set({ ... })".
        type [<AllowNullLiteral>] ParagraphCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.ParagraphData> option with get, set

        /// An interface for updating data on the Image object, for use in "image.set({ ... })".
        type [<AllowNullLiteral>] ImageUpdateData =
            /// Gets or sets the description of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract description: string option with get, set
            /// Gets or sets the height of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract height: float option with get, set
            /// Gets or sets the hyperlink of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract hyperlink: string option with get, set
            /// Gets or sets the width of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract width: float option with get, set

        /// An interface for updating data on the Table object, for use in "table.set({ ... })".
        type [<AllowNullLiteral>] TableUpdateData =
            /// Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract borderVisible: bool option with get, set

        /// An interface for updating data on the TableRowCollection object, for use in "tableRowCollection.set({ ... })".
        type [<AllowNullLiteral>] TableRowCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.TableRowData> option with get, set

        /// An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })".
        type [<AllowNullLiteral>] TableCellUpdateData =
            /// Gets and sets the shading color of the cell
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract shadingColor: string option with get, set

        /// An interface for updating data on the TableCellCollection object, for use in "tableCellCollection.set({ ... })".
        type [<AllowNullLiteral>] TableCellCollectionUpdateData =
            abstract items: ResizeArray<OneNote.Interfaces.TableCellData> option with get, set

        /// An interface describing the data returned by calling "application.toJSON()".
        type [<AllowNullLiteral>] ApplicationData =
            /// Gets the collection of notebooks that are open in the OneNote application instance. In OneNote on the web, only one notebook at a time is open in the application instance. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebooks: ResizeArray<OneNote.Interfaces.NotebookData> option with get, set

        /// An interface describing the data returned by calling "inkAnalysis.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisData =
            /// Gets the parent page object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract page: OneNote.Interfaces.PageData option with get, set
            /// Gets the ID of the InkAnalysis object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "inkAnalysisParagraph.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisParagraphData =
            /// Reference to the parent InkAnalysisPage. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysis: OneNote.Interfaces.InkAnalysisData option with get, set
            /// Gets the ink analysis lines in this ink analysis paragraph. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract lines: ResizeArray<OneNote.Interfaces.InkAnalysisLineData> option with get, set
            /// Gets the ID of the InkAnalysisParagraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "inkAnalysisParagraphCollection.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisParagraphCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisParagraphData> option with get, set

        /// An interface describing the data returned by calling "inkAnalysisLine.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisLineData =
            /// Reference to the parent InkAnalysisParagraph. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.InkAnalysisParagraphData option with get, set
            /// Gets the ink analysis words in this ink analysis line. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract words: ResizeArray<OneNote.Interfaces.InkAnalysisWordData> option with get, set
            /// Gets the ID of the InkAnalysisLine object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "inkAnalysisLineCollection.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisLineCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisLineData> option with get, set

        /// An interface describing the data returned by calling "inkAnalysisWord.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisWordData =
            /// Reference to the parent InkAnalysisLine. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract line: OneNote.Interfaces.InkAnalysisLineData option with get, set
            /// Gets the ID of the InkAnalysisWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// The id of the recognized language in this inkAnalysisWord. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: string option with get, set
            /// Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract strokePointers: ResizeArray<OneNote.InkStrokePointer> option with get, set
            /// The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: ResizeArray<string> option with get, set

        /// An interface describing the data returned by calling "inkAnalysisWordCollection.toJSON()".
        type [<AllowNullLiteral>] InkAnalysisWordCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.InkAnalysisWordData> option with get, set

        /// An interface describing the data returned by calling "floatingInk.toJSON()".
        type [<AllowNullLiteral>] FloatingInkData =
            /// Gets the strokes of the FloatingInk object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkStrokes: ResizeArray<OneNote.Interfaces.InkStrokeData> option with get, set
            /// Gets the ID of the FloatingInk object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "inkStroke.toJSON()".
        type [<AllowNullLiteral>] InkStrokeData =
            /// Gets the ID of the InkStroke object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract floatingInk: OneNote.Interfaces.FloatingInkData option with get, set
            /// Gets the ID of the InkStroke object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "inkStrokeCollection.toJSON()".
        type [<AllowNullLiteral>] InkStrokeCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.InkStrokeData> option with get, set

        /// An interface describing the data returned by calling "inkWord.toJSON()".
        type [<AllowNullLiteral>] InkWordData =
            /// Gets the ID of the InkWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// The id of the recognized language in this ink word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: string option with get, set
            /// The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: ResizeArray<string> option with get, set

        /// An interface describing the data returned by calling "inkWordCollection.toJSON()".
        type [<AllowNullLiteral>] InkWordCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.InkWordData> option with get, set

        /// An interface describing the data returned by calling "notebook.toJSON()".
        type [<AllowNullLiteral>] NotebookData =
            /// The section groups in the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: ResizeArray<OneNote.Interfaces.SectionGroupData> option with get, set
            /// The the sections of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: ResizeArray<OneNote.Interfaces.SectionData> option with get, set
            /// The url of the site that this notebook is located. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract baseUrl: string option with get, set
            /// The client url of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: string option with get, set
            /// Gets the ID of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isVirtual: bool option with get, set
            /// Gets the name of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: string option with get, set

        /// An interface describing the data returned by calling "notebookCollection.toJSON()".
        type [<AllowNullLiteral>] NotebookCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.NotebookData> option with get, set

        /// An interface describing the data returned by calling "sectionGroup.toJSON()".
        type [<AllowNullLiteral>] SectionGroupData =
            /// The collection of section groups in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: ResizeArray<OneNote.Interfaces.SectionGroupData> option with get, set
            /// The collection of sections in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: ResizeArray<OneNote.Interfaces.SectionData> option with get, set
            /// The client url of the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: string option with get, set
            /// Gets the ID of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the name of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: string option with get, set

        /// An interface describing the data returned by calling "sectionGroupCollection.toJSON()".
        type [<AllowNullLiteral>] SectionGroupCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.SectionGroupData> option with get, set

        /// An interface describing the data returned by calling "section.toJSON()".
        type [<AllowNullLiteral>] SectionData =
            /// The collection of pages in the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pages: ResizeArray<OneNote.Interfaces.PageData> option with get, set
            /// The client url of the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: string option with get, set
            /// Gets the ID of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// True if this section is encrypted with a password. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isEncrypted: bool option with get, set
            /// True if this section is locked. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isLocked: bool option with get, set
            /// Gets the name of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: string option with get, set
            /// The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: string option with get, set

        /// An interface describing the data returned by calling "sectionCollection.toJSON()".
        type [<AllowNullLiteral>] SectionCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.SectionData> option with get, set

        /// An interface describing the data returned by calling "page.toJSON()".
        type [<AllowNullLiteral>] PageData =
            /// The collection of PageContent objects on the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract contents: ResizeArray<OneNote.Interfaces.PageContentData> option with get, set
            /// Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysisOrNull: OneNote.Interfaces.InkAnalysisData option with get, set
            /// Gets the ClassNotebookPageSource to the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract classNotebookPageSource: string option with get, set
            /// The client url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: string option with get, set
            /// Gets the ID of the page. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets or sets the indentation level of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageLevel: float option with get, set
            /// Gets or sets the title of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract title: string option with get, set
            /// The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: string option with get, set

        /// An interface describing the data returned by calling "pageCollection.toJSON()".
        type [<AllowNullLiteral>] PageCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.PageData> option with get, set

        /// An interface describing the data returned by calling "pageContent.toJSON()".
        type [<AllowNullLiteral>] PageContentData =
            /// Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageData option with get, set
            /// Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ink: OneNote.Interfaces.FloatingInkData option with get, set
            /// Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract outline: OneNote.Interfaces.OutlineData option with get, set
            /// Gets the ID of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets or sets the left (X-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract left: float option with get, set
            /// Gets or sets the top (Y-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract top: float option with get, set
            /// Gets the type of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: U2<OneNote.PageContentType, string> option with get, set

        /// An interface describing the data returned by calling "pageContentCollection.toJSON()".
        type [<AllowNullLiteral>] PageContentCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.PageContentData> option with get, set

        /// An interface describing the data returned by calling "outline.toJSON()".
        type [<AllowNullLiteral>] OutlineData =
            /// Gets the collection of Paragraph objects in the Outline. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: ResizeArray<OneNote.Interfaces.ParagraphData> option with get, set
            /// Gets the ID of the Outline object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set

        /// An interface describing the data returned by calling "paragraph.toJSON()".
        type [<AllowNullLiteral>] ParagraphData =
            /// Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageData option with get, set
            /// Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkWords: ResizeArray<OneNote.Interfaces.InkWordData> option with get, set
            /// The collection of paragraphs under this paragraph. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: ResizeArray<OneNote.Interfaces.ParagraphData> option with get, set
            /// Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract richText: OneNote.Interfaces.RichTextData option with get, set
            /// Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract table: OneNote.Interfaces.TableData option with get, set
            /// Gets the ID of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the type of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: U2<OneNote.ParagraphType, string> option with get, set

        /// An interface describing the data returned by calling "paragraphCollection.toJSON()".
        type [<AllowNullLiteral>] ParagraphCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.ParagraphData> option with get, set

        /// An interface describing the data returned by calling "noteTag.toJSON()".
        type [<AllowNullLiteral>] NoteTagData =
            /// Gets the Id of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the status of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract status: U2<OneNote.NoteTagStatus, string> option with get, set
            /// Gets the type of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: U2<OneNote.NoteTagType, string> option with get, set

        /// An interface describing the data returned by calling "richText.toJSON()".
        type [<AllowNullLiteral>] RichTextData =
            /// Gets the ID of the RichText object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// The language id of the text. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: string option with get, set
            /// Gets the text content of the RichText object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract text: string option with get, set

        /// An interface describing the data returned by calling "image.toJSON()".
        type [<AllowNullLiteral>] ImageData =
            /// Gets or sets the description of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract description: string option with get, set
            /// Gets or sets the height of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract height: float option with get, set
            /// Gets or sets the hyperlink of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract hyperlink: string option with get, set
            /// Gets the ID of the Image object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ocrData: OneNote.ImageOcrData option with get, set
            /// Gets or sets the width of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract width: float option with get, set

        /// An interface describing the data returned by calling "table.toJSON()".
        type [<AllowNullLiteral>] TableData =
            /// Gets all of the table rows. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rows: ResizeArray<OneNote.Interfaces.TableRowData> option with get, set
            /// Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract borderVisible: bool option with get, set
            /// Gets the number of columns in the table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract columnCount: float option with get, set
            /// Gets the ID of the table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the number of rows in the table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowCount: float option with get, set

        /// An interface describing the data returned by calling "tableRow.toJSON()".
        type [<AllowNullLiteral>] TableRowData =
            /// Gets the cells in the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cells: ResizeArray<OneNote.Interfaces.TableCellData> option with get, set
            /// Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellCount: float option with get, set
            /// Gets the ID of the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: float option with get, set

        /// An interface describing the data returned by calling "tableRowCollection.toJSON()".
        type [<AllowNullLiteral>] TableRowCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.TableRowData> option with get, set

        /// An interface describing the data returned by calling "tableCell.toJSON()".
        type [<AllowNullLiteral>] TableCellData =
            /// Gets the collection of Paragraph objects in the TableCell. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: ResizeArray<OneNote.Interfaces.ParagraphData> option with get, set
            /// Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellIndex: float option with get, set
            /// Gets the ID of the cell. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: string option with get, set
            /// Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: float option with get, set
            /// Gets and sets the shading color of the cell
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract shadingColor: string option with get, set

        /// An interface describing the data returned by calling "tableCellCollection.toJSON()".
        type [<AllowNullLiteral>] TableCellCollectionData =
            abstract items: ResizeArray<OneNote.Interfaces.TableCellData> option with get, set

        /// Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] ApplicationLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the collection of notebooks that are open in the OneNote application instance. In OneNote on the web, only one notebook at a time is open in the application instance.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebooks: OneNote.Interfaces.NotebookCollectionLoadOptions option with get, set

        /// Represents ink analysis data for a given set of ink strokes.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the parent page object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract page: OneNote.Interfaces.PageLoadOptions option with get, set
            /// Gets the ID of the InkAnalysis object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents ink analysis data for an identified paragraph formed by ink strokes.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisParagraphLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Reference to the parent InkAnalysisPage.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysis: OneNote.Interfaces.InkAnalysisLoadOptions option with get, set
            /// Gets the ink analysis lines in this ink analysis paragraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract lines: OneNote.Interfaces.InkAnalysisLineCollectionLoadOptions option with get, set
            /// Gets the ID of the InkAnalysisParagraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents a collection of InkAnalysisParagraph objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisParagraphCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Reference to the parent InkAnalysisPage.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysis: OneNote.Interfaces.InkAnalysisLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ink analysis lines in this ink analysis paragraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract lines: OneNote.Interfaces.InkAnalysisLineCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkAnalysisParagraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents ink analysis data for an identified text line formed by ink strokes.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisLineLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Reference to the parent InkAnalysisParagraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.InkAnalysisParagraphLoadOptions option with get, set
            /// Gets the ink analysis words in this ink analysis line.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract words: OneNote.Interfaces.InkAnalysisWordCollectionLoadOptions option with get, set
            /// Gets the ID of the InkAnalysisLine object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents a collection of InkAnalysisLine objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisLineCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Reference to the parent InkAnalysisParagraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.InkAnalysisParagraphLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ink analysis words in this ink analysis line.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract words: OneNote.Interfaces.InkAnalysisWordCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkAnalysisLine object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents ink analysis data for an identified word formed by ink strokes.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisWordLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Reference to the parent InkAnalysisLine.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract line: OneNote.Interfaces.InkAnalysisLineLoadOptions option with get, set
            /// Gets the ID of the InkAnalysisWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// The id of the recognized language in this inkAnalysisWord. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: bool option with get, set
            /// Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract strokePointers: bool option with get, set
            /// The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: bool option with get, set

        /// Represents a collection of InkAnalysisWord objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkAnalysisWordCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Reference to the parent InkAnalysisLine.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract line: OneNote.Interfaces.InkAnalysisLineLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkAnalysisWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: The id of the recognized language in this inkAnalysisWord. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: bool option with get, set
            /// For EACH ITEM in the collection: Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract strokePointers: bool option with get, set
            /// For EACH ITEM in the collection: The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: bool option with get, set

        /// Represents a group of ink strokes.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] FloatingInkLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the strokes of the FloatingInk object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkStrokes: OneNote.Interfaces.InkStrokeCollectionLoadOptions option with get, set
            /// Gets the PageContent parent of the FloatingInk object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageContent: OneNote.Interfaces.PageContentLoadOptions option with get, set
            /// Gets the ID of the FloatingInk object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents a single stroke of ink.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkStrokeLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the ID of the InkStroke object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract floatingInk: OneNote.Interfaces.FloatingInkLoadOptions option with get, set
            /// Gets the ID of the InkStroke object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// Represents a collection of InkStroke objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkStrokeCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkStroke object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract floatingInk: OneNote.Interfaces.FloatingInkLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkStroke object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// A container for the ink in a word in a paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkWordLoadOptions =
            abstract ``$all``: bool option with get, set
            /// The parent paragraph containing the ink word.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets the ID of the InkWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// The id of the recognized language in this ink word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: bool option with get, set
            /// The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: bool option with get, set

        /// Represents a collection of InkWord objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] InkWordCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: The parent paragraph containing the ink word.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the InkWord object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: The id of the recognized language in this ink word. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: bool option with get, set
            /// For EACH ITEM in the collection: The words that were recognized in this ink word, in order of likelihood. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract wordAlternates: bool option with get, set

        /// Represents a OneNote notebook. Notebooks contain section groups and sections.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] NotebookLoadOptions =
            abstract ``$all``: bool option with get, set
            /// The section groups in the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: OneNote.Interfaces.SectionGroupCollectionLoadOptions option with get, set
            /// The the sections of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: OneNote.Interfaces.SectionCollectionLoadOptions option with get, set
            /// The url of the site that this notebook is located. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract baseUrl: bool option with get, set
            /// The client url of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// Gets the ID of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isVirtual: bool option with get, set
            /// Gets the name of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set

        /// Represents a collection of notebooks.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] NotebookCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: The section groups in the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: OneNote.Interfaces.SectionGroupCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: The the sections of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: OneNote.Interfaces.SectionCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: The url of the site that this notebook is located. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract baseUrl: bool option with get, set
            /// For EACH ITEM in the collection: The client url of the notebook. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isVirtual: bool option with get, set
            /// For EACH ITEM in the collection: Gets the name of the notebook. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set

        /// Represents a OneNote section group. Section groups can contain sections and other section groups.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] SectionGroupLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the notebook that contains the section group.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebook: OneNote.Interfaces.NotebookLoadOptions option with get, set
            /// Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroup: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroupOrNull: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// The collection of section groups in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: OneNote.Interfaces.SectionGroupCollectionLoadOptions option with get, set
            /// The collection of sections in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: OneNote.Interfaces.SectionCollectionLoadOptions option with get, set
            /// The client url of the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// Gets the ID of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the name of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set

        /// Represents a collection of section groups.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] SectionGroupCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the notebook that contains the section group.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebook: OneNote.Interfaces.NotebookLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroup: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroupOrNull: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// For EACH ITEM in the collection: The collection of section groups in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sectionGroups: OneNote.Interfaces.SectionGroupCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: The collection of sections in the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract sections: OneNote.Interfaces.SectionCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: The client url of the section group. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets the name of the section group. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set

        /// Represents a OneNote section. Sections can contain pages.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] SectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the notebook that contains the section.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebook: OneNote.Interfaces.NotebookLoadOptions option with get, set
            /// The collection of pages in the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pages: OneNote.Interfaces.PageCollectionLoadOptions option with get, set
            /// Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroup: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// Gets the section group that contains the section. Returns null if the section is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroupOrNull: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// The client url of the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// Gets the ID of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// True if this section is encrypted with a password. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isEncrypted: bool option with get, set
            /// True if this section is locked. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isLocked: bool option with get, set
            /// Gets the name of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set
            /// The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: bool option with get, set

        /// Represents a collection of sections.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] SectionCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the notebook that contains the section.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract notebook: OneNote.Interfaces.NotebookLoadOptions option with get, set
            /// For EACH ITEM in the collection: The collection of pages in the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pages: OneNote.Interfaces.PageCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroup: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the section group that contains the section. Returns null if the section is a direct child of the notebook.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSectionGroupOrNull: OneNote.Interfaces.SectionGroupLoadOptions option with get, set
            /// For EACH ITEM in the collection: The client url of the section. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: True if this section is encrypted with a password. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isEncrypted: bool option with get, set
            /// For EACH ITEM in the collection: True if this section is locked. Read only
            /// 
            /// [Api set: OneNoteApi 1.2]
            abstract isLocked: bool option with get, set
            /// For EACH ITEM in the collection: Gets the name of the section. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract name: bool option with get, set
            /// For EACH ITEM in the collection: The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: bool option with get, set

        /// Represents a OneNote page.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] PageLoadOptions =
            abstract ``$all``: bool option with get, set
            /// The collection of PageContent objects on the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract contents: OneNote.Interfaces.PageContentCollectionLoadOptions option with get, set
            /// Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysisOrNull: OneNote.Interfaces.InkAnalysisLoadOptions option with get, set
            /// Gets the section that contains the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSection: OneNote.Interfaces.SectionLoadOptions option with get, set
            /// Gets the ClassNotebookPageSource to the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract classNotebookPageSource: bool option with get, set
            /// The client url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// Gets the ID of the page. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets or sets the indentation level of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageLevel: bool option with get, set
            /// Gets or sets the title of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract title: bool option with get, set
            /// The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: bool option with get, set

        /// Represents a collection of pages.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] PageCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: The collection of PageContent objects on the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract contents: OneNote.Interfaces.PageContentCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkAnalysisOrNull: OneNote.Interfaces.InkAnalysisLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the section that contains the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentSection: OneNote.Interfaces.SectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ClassNotebookPageSource to the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract classNotebookPageSource: bool option with get, set
            /// For EACH ITEM in the collection: The client url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract clientUrl: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the page. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the indentation level of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageLevel: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the title of the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract title: bool option with get, set
            /// For EACH ITEM in the collection: The web url of the page. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract webUrl: bool option with get, set

        /// Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] PageContentLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageLoadOptions option with get, set
            /// Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ink: OneNote.Interfaces.FloatingInkLoadOptions option with get, set
            /// Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract outline: OneNote.Interfaces.OutlineLoadOptions option with get, set
            /// Gets the page that contains the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentPage: OneNote.Interfaces.PageLoadOptions option with get, set
            /// Gets the ID of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets or sets the left (X-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract left: bool option with get, set
            /// Gets or sets the top (Y-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract top: bool option with get, set
            /// Gets the type of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: bool option with get, set

        /// Represents the contents of a page, as a collection of PageContent objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] PageContentCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ink: OneNote.Interfaces.FloatingInkLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract outline: OneNote.Interfaces.OutlineLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the page that contains the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentPage: OneNote.Interfaces.PageLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the left (X-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract left: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the top (Y-axis) position of the PageContent object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract top: bool option with get, set
            /// For EACH ITEM in the collection: Gets the type of the PageContent object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: bool option with get, set

        /// Represents a container for Paragraph objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] OutlineLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageContent: OneNote.Interfaces.PageContentLoadOptions option with get, set
            /// Gets the collection of Paragraph objects in the Outline.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: OneNote.Interfaces.ParagraphCollectionLoadOptions option with get, set
            /// Gets the ID of the Outline object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set

        /// A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] ParagraphLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageLoadOptions option with get, set
            /// Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkWords: OneNote.Interfaces.InkWordCollectionLoadOptions option with get, set
            /// Gets the Outline object that contains the Paragraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract outline: OneNote.Interfaces.OutlineLoadOptions option with get, set
            /// The collection of paragraphs under this paragraph. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: OneNote.Interfaces.ParagraphCollectionLoadOptions option with get, set
            /// Gets the parent paragraph object. Throws if a parent paragraph does not exist.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentParagraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets the parent paragraph object. Returns null if a parent paragraph does not exist.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentParagraphOrNull: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTableCell: OneNote.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTableCellOrNull: OneNote.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract richText: OneNote.Interfaces.RichTextLoadOptions option with get, set
            /// Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract table: OneNote.Interfaces.TableLoadOptions option with get, set
            /// Gets the ID of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the type of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: bool option with get, set

        /// Represents a collection of Paragraph objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] ParagraphCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract image: OneNote.Interfaces.ImageLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract inkWords: OneNote.Interfaces.InkWordCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the Outline object that contains the Paragraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract outline: OneNote.Interfaces.OutlineLoadOptions option with get, set
            /// For EACH ITEM in the collection: The collection of paragraphs under this paragraph. Read only
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: OneNote.Interfaces.ParagraphCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent paragraph object. Throws if a parent paragraph does not exist.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentParagraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent paragraph object. Returns null if a parent paragraph does not exist.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentParagraphOrNull: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTableCell: OneNote.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTableCellOrNull: OneNote.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract richText: OneNote.Interfaces.RichTextLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract table: OneNote.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets the type of the Paragraph object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: bool option with get, set

        /// A container for the NoteTag in a paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] NoteTagLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the Id of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the status of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract status: bool option with get, set
            /// Gets the type of the NoteTag object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ``type``: bool option with get, set

        /// Represents a RichText object in a Paragraph.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] RichTextLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the Paragraph object that contains the RichText object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets the ID of the RichText object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// The language id of the text. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract languageId: bool option with get, set
            /// Gets the text content of the RichText object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract text: bool option with get, set

        /// Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] ImageLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract pageContent: OneNote.Interfaces.PageContentLoadOptions option with get, set
            /// Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets or sets the description of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract description: bool option with get, set
            /// Gets or sets the height of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract height: bool option with get, set
            /// Gets or sets the hyperlink of the Image.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract hyperlink: bool option with get, set
            /// Gets the ID of the Image object. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract ocrData: bool option with get, set
            /// Gets or sets the width of the Image layout.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract width: bool option with get, set

        /// Represents a table in a OneNote page.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] TableLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the Paragraph object that contains the Table object.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraph: OneNote.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets all of the table rows.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rows: OneNote.Interfaces.TableRowCollectionLoadOptions option with get, set
            /// Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract borderVisible: bool option with get, set
            /// Gets the number of columns in the table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract columnCount: bool option with get, set
            /// Gets the ID of the table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the number of rows in the table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowCount: bool option with get, set

        /// Represents a row in a table.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] TableRowLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the cells in the row.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cells: OneNote.Interfaces.TableCellCollectionLoadOptions option with get, set
            /// Gets the parent table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTable: OneNote.Interfaces.TableLoadOptions option with get, set
            /// Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellCount: bool option with get, set
            /// Gets the ID of the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: bool option with get, set

        /// Contains a collection of TableRow objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] TableRowCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the cells in the row.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cells: OneNote.Interfaces.TableCellCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent table.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentTable: OneNote.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellCount: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: bool option with get, set

        /// Represents a cell in a OneNote table.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] TableCellLoadOptions =
            abstract ``$all``: bool option with get, set
            /// Gets the collection of Paragraph objects in the TableCell.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: OneNote.Interfaces.ParagraphCollectionLoadOptions option with get, set
            /// Gets the parent row of the cell.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentRow: OneNote.Interfaces.TableRowLoadOptions option with get, set
            /// Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellIndex: bool option with get, set
            /// Gets the ID of the cell. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: bool option with get, set
            /// Gets and sets the shading color of the cell
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract shadingColor: bool option with get, set

        /// Contains a collection of TableCell objects.
        /// 
        /// [Api set: OneNoteApi 1.1]
        type [<AllowNullLiteral>] TableCellCollectionLoadOptions =
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the collection of Paragraph objects in the TableCell.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract paragraphs: OneNote.Interfaces.ParagraphCollectionLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent row of the cell.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract parentRow: OneNote.Interfaces.TableRowLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract cellIndex: bool option with get, set
            /// For EACH ITEM in the collection: Gets the ID of the cell. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract rowIndex: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the shading color of the cell
            /// 
            /// [Api set: OneNoteApi 1.1]
            abstract shadingColor: bool option with get, set

    type [<AllowNullLiteral>] RequestContext =
        inherit OfficeCore.RequestContext
        abstract application: Application

    type [<AllowNullLiteral>] RequestContextStatic =
        [<Emit "new $0($1...)">] abstract Create: ?url: string -> RequestContext