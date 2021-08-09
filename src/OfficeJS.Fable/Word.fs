namespace OfficeJS.Fable

open System
open Fable.Core
open Fable.Core.JS
open Browser.Types


module rec Word =

    type [<AllowNullLiteral>] IExports =
        abstract Application: ApplicationStatic
        abstract Body: BodyStatic
        abstract ContentControl: ContentControlStatic
        abstract ContentControlCollection: ContentControlCollectionStatic
        abstract CustomProperty: CustomPropertyStatic
        abstract CustomPropertyCollection: CustomPropertyCollectionStatic
        abstract Document: DocumentStatic
        abstract DocumentCreated: DocumentCreatedStatic
        abstract DocumentProperties: DocumentPropertiesStatic
        abstract Font: FontStatic
        abstract InlinePicture: InlinePictureStatic
        abstract InlinePictureCollection: InlinePictureCollectionStatic
        abstract List: ListStatic
        abstract ListCollection: ListCollectionStatic
        abstract ListItem: ListItemStatic
        abstract Paragraph: ParagraphStatic
        abstract ParagraphCollection: ParagraphCollectionStatic
        abstract Range: RangeStatic
        abstract RangeCollection: RangeCollectionStatic
        abstract SearchOptions: SearchOptionsStatic
        abstract Section: SectionStatic
        abstract SectionCollection: SectionCollectionStatic
        abstract Table: TableStatic
        abstract TableCollection: TableCollectionStatic
        abstract TableRow: TableRowStatic
        abstract TableRowCollection: TableRowCollectionStatic
        abstract TableCell: TableCellStatic
        abstract TableCellCollection: TableCellCollectionStatic
        abstract TableBorder: TableBorderStatic
        abstract RequestContext: RequestContextStatic
        /// <summary>Executes a batch script that performs actions on the Word object model, using the RequestContext of previously created API objects.</summary>
        /// <param name="objects">- An array of previously created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".</param>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.</param>
        abstract run: objects: ResizeArray<OfficeExtension.ClientObject> * batch: (Word.RequestContext -> Promise<'T>) -> Promise<'T>
        /// <summary>Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</summary>
        /// <param name="object">- A previously created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".</param>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.</param>
        abstract run: ``object``: OfficeExtension.ClientObject * batch: (Word.RequestContext -> Promise<'T>) -> Promise<'T>
        /// <summary>Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.</summary>
        /// <param name="batch">- A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.</param>
        abstract run: batch: (Word.RequestContext -> Promise<'T>) -> Promise<'T>

    /// Represents the application object.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] Application =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// <summary>Creates a new document by using an optional base64 encoded .docx file.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="base64File">Optional. The base64 encoded .docx file. The default value is null.</param>
        abstract createDocument: ?base64File: string -> Word.DocumentCreated
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> ApplicationToJSONReturn

    type [<AllowNullLiteral>] ApplicationToJSONReturn =
        [<Emit "$0[$1]{{=$2}}">] abstract Item: key: string -> string with get, set

    /// Represents the application object.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ApplicationStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Application
        /// Create a new instance of Word.Application object
        abstract newObject: context: OfficeExtension.ClientRequestContext -> Word.Application

    /// Represents the body of a document or a section.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Body =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of rich text content control objects in the body. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract font: Word.Font
        /// Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract inlinePictures: Word.InlinePictureCollection
        /// Gets the collection of list objects in the body. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lists: Word.ListCollection
        /// Gets the collection of paragraph objects in the body. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract paragraphs: Word.ParagraphCollection
        /// Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an error if there isn't a parent body. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBody: Word.Body
        /// Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBodyOrNullObject: Word.Body
        /// Gets the content control that contains the body. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the parent section of the body. Throws an error if there isn't a parent section. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentSection: Word.Section
        /// Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentSectionOrNullObject: Word.Section
        /// Gets the collection of table objects in the body. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract tables: Word.TableCollection
        /// Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
        /// 
        /// [Api set: WordApi 1.1]
        abstract style: string with get, set
        /// Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBuiltIn: U2<Word.Style, string> with get, set
        /// Gets the text of the body. Use the insertText method to insert text. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract text: string
        /// Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract ``type``: U2<Word.BodyType, string>
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.BodyUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Body -> unit
        /// Clears the contents of the body object. The user can perform the undo operation on the cleared content.
        /// 
        /// [Api set: WordApi 1.1]
        abstract clear: unit -> unit
        /// Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getHtml: unit -> OfficeExtension.ClientResult<string>
        /// Gets the OOXML (Office Open XML) representation of the body object.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getOoxml: unit -> OfficeExtension.ClientResult<string>
        /// <summary>Gets the whole body, or the starting or ending point of the body, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Gets the whole body, or the starting or ending point of the body, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: BodyGetRangeRangeLocation -> Word.Range
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add to the body.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertBreak: breakType: Word.BreakType * insertLocation: Word.InsertLocation -> unit
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add to the body.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertBreak: breakType: BodyInsertBreakBreakType * insertLocation: BodyInsertBreakInsertLocation -> unit
        /// Wraps the body object with a Rich Text content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract insertContentControl: unit -> Word.ContentControl
        /// <summary>Inserts a document into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts a document into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: BodyInsertFileFromBase64InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in the document.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertHtml: html: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in the document.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertHtml: html: string * insertLocation: BodyInsertHtmlInsertLocation -> Word.Range
        /// <summary>Inserts a picture into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the body.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: Word.InsertLocation -> Word.InlinePicture
        /// <summary>Inserts a picture into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the body.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: BodyInsertInlinePictureFromBase64InsertLocation -> Word.InlinePicture
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: BodyInsertOoxmlInsertLocation -> Word.Range
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: BodyInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: Word.InsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: BodyInsertTableInsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts text into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertText: text: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts text into the body at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertText: text: string * insertLocation: BodyInsertTextInsertLocation -> Word.Range
        /// <summary>Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="searchText">Required. The search text. Can be a maximum of 255 characters.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects the body and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects the body and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: BodySelectSelectionMode -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.BodyLoadOptions -> Word.Body
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Body
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: BodyLoadPropertyNamesAndPaths -> Word.Body
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Body
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Body
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Body object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.BodyData

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertBreakBreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertBreakInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertFileFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertHtmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertInlinePictureFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertOoxmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertTableInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyInsertTextInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] BodySelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<AllowNullLiteral>] BodyLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents the body of a document or a section.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] BodyStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Body

    /// Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ContentControl =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of content control objects in the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract font: Word.Font
        /// Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract inlinePictures: Word.InlinePictureCollection
        /// Gets the collection of list objects in the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lists: Word.ListCollection
        /// Get the collection of paragraph objects in the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract paragraphs: Word.ParagraphCollection
        /// Gets the parent body of the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBody: Word.Body
        /// Gets the content control that contains the content control. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the table that contains the content control. Throws an error if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCell: Word.TableCell
        /// Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCellOrNullObject: Word.TableCell
        /// Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableOrNullObject: Word.Table
        /// Gets the collection of table objects in the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract tables: Word.TableCollection
        /// Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
        /// 
        /// [Api set: WordApi 1.1]
        abstract appearance: U2<Word.ContentControlAppearance, string> with get, set
        /// Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
        /// 
        /// [Api set: WordApi 1.1]
        abstract cannotDelete: bool with get, set
        /// Gets or sets a value that indicates whether the user can edit the contents of the content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract cannotEdit: bool with get, set
        /// Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
        /// 
        /// [Api set: WordApi 1.1]
        abstract color: string with get, set
        /// Gets an integer that represents the content control identifier. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract id: float
        /// Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
        /// 
        /// **Note**: The set operation for this property is not supported in Word on the web.
        /// 
        /// [Api set: WordApi 1.1]
        abstract placeholderText: string with get, set
        /// Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
        /// 
        /// [Api set: WordApi 1.1]
        abstract removeWhenEdited: bool with get, set
        /// Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
        /// 
        /// [Api set: WordApi 1.1]
        abstract style: string with get, set
        /// Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBuiltIn: U2<Word.Style, string> with get, set
        /// Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract subtype: U2<Word.ContentControlType, string>
        /// Gets or sets a tag to identify a content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract tag: string with get, set
        /// Gets the text of the content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract text: string
        /// Gets or sets the title for a content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract title: string with get, set
        /// Gets the content control type. Only rich text content controls are supported currently. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract ``type``: U2<Word.ContentControlType, string>
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.ContentControlUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.ContentControl -> unit
        /// Clears the contents of the content control. The user can perform the undo operation on the cleared content.
        /// 
        /// [Api set: WordApi 1.1]
        abstract clear: unit -> unit
        /// <summary>Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="keepContent">Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.</param>
        abstract delete: keepContent: bool -> unit
        /// Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getHtml: unit -> OfficeExtension.ClientResult<string>
        /// Gets the Office Open XML (OOXML) representation of the content control object.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getOoxml: unit -> OfficeExtension.ClientResult<string>
        /// <summary>Gets the whole content control, or the starting or ending point of the content control, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Gets the whole content control, or the starting or ending point of the content control, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: ContentControlGetRangeRangeLocation -> Word.Range
        /// <summary>Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract getTextRanges: endingMarks: ResizeArray<string> * ?trimSpacing: bool -> Word.RangeCollection
        /// <summary>Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. Type of break.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertBreak: breakType: Word.BreakType * insertLocation: Word.InsertLocation -> unit
        /// <summary>Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. Type of break.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertBreak: breakType: ContentControlInsertBreakBreakType * insertLocation: ContentControlInsertBreakInsertLocation -> unit
        /// <summary>Inserts a document into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts a document into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: ContentControlInsertFileFromBase64InsertLocation -> Word.Range
        /// <summary>Inserts HTML into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertHtml: html: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts HTML into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertHtml: html: string * insertLocation: ContentControlInsertHtmlInsertLocation -> Word.Range
        /// <summary>Inserts an inline picture into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: Word.InsertLocation -> Word.InlinePicture
        /// <summary>Inserts an inline picture into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: ContentControlInsertInlinePictureFromBase64InsertLocation -> Word.InlinePicture
        /// <summary>Inserts OOXML into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertOoxml: ooxml: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts OOXML into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertOoxml: ooxml: string * insertLocation: ContentControlInsertOoxmlInsertLocation -> Word.Range
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: ContentControlInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts a table with the specified number of rows and columns into, or next to, a content control.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: Word.InsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts a table with the specified number of rows and columns into, or next to, a content control.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: ContentControlInsertTableInsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts text into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. The text to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertText: text: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts text into the content control at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. The text to be inserted in to the content control.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
        abstract insertText: text: string * insertLocation: ContentControlInsertTextInsertLocation -> Word.Range
        /// <summary>Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="searchText">Required. The search text.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects the content control. This causes Word to scroll to the selection.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects the content control. This causes Word to scroll to the selection.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: ContentControlSelectSelectionMode -> unit
        /// <summary>Splits the content control into child ranges by using delimiters.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="delimiters">Required. The delimiters as an array of strings.</param>
        /// <param name="multiParagraphs">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
        /// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract split: delimiters: ResizeArray<string> * ?multiParagraphs: bool * ?trimDelimiters: bool * ?trimSpacing: bool -> Word.RangeCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.ContentControlLoadOptions -> Word.ContentControl
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.ContentControl
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: ContentControlLoadPropertyNamesAndPaths -> Word.ContentControl
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.ContentControl
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.ContentControl
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.ContentControl object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.ContentControlData

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertBreakBreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertBreakInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertFileFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertHtmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertInlinePictureFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertOoxmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertTableInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlInsertTextInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<AllowNullLiteral>] ContentControlLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ContentControlStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ContentControl

    /// Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ContentControlCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.ContentControl>
        /// <summary>Gets a content control by its identifier. Throws an error if there isn't a content control with the identifier in this collection.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="id">Required. A content control identifier.</param>
        abstract getById: id: float -> Word.ContentControl
        /// <summary>Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="id">Required. A content control identifier.</param>
        abstract getByIdOrNullObject: id: float -> Word.ContentControl
        /// <summary>Gets the content controls that have the specified tag.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="tag">Required. A tag set on a content control.</param>
        abstract getByTag: tag: string -> Word.ContentControlCollection
        /// <summary>Gets the content controls that have the specified title.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="title">Required. The title of a content control.</param>
        abstract getByTitle: title: string -> Word.ContentControlCollection
        /// <summary>Gets the content controls that have the specified types and/or subtypes.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="types">Required. An array of content control types and/or subtypes.</param>
        abstract getByTypes: types: ResizeArray<Word.ContentControlType> -> Word.ContentControlCollection
        /// Gets the first content control in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.ContentControl
        /// Gets the first content control in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.ContentControl
        /// <summary>Gets a content control by its index in the collection.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="index">The index.</param>
        abstract getItem: index: float -> Word.ContentControl
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.ContentControlCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.ContentControlCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.ContentControlCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.ContentControlCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.ContentControlCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.ContentControlCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.ContentControlCollectionData

    /// Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ContentControlCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ContentControlCollection

    /// Represents a custom property.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] CustomProperty =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the key of the custom property. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract key: string
        /// Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract ``type``: U2<Word.DocumentPropertyType, string>
        /// Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
        /// 
        /// [Api set: WordApi 1.3]
        abstract value: obj option with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.CustomPropertyUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.CustomProperty -> unit
        /// Deletes the custom property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract delete: unit -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.CustomPropertyLoadOptions -> Word.CustomProperty
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.CustomProperty
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: CustomPropertyLoadPropertyNamesAndPaths -> Word.CustomProperty
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.CustomProperty
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.CustomProperty
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.CustomProperty object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.CustomPropertyData

    type [<AllowNullLiteral>] CustomPropertyLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a custom property.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] CustomPropertyStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> CustomProperty

    /// Contains the collection of {@link Word.CustomProperty} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] CustomPropertyCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.CustomProperty>
        /// <summary>Creates a new or sets an existing custom property.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="key">Required. The custom property's key, which is case-insensitive.</param>
        /// <param name="value">Required. The custom property's value.</param>
        abstract add: key: string * value: obj option -> Word.CustomProperty
        /// Deletes all custom properties in this collection.
        /// 
        /// [Api set: WordApi 1.3]
        abstract deleteAll: unit -> unit
        /// Gets the count of custom properties.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getCount: unit -> OfficeExtension.ClientResult<float>
        /// <summary>Gets a custom property object by its key, which is case-insensitive. Throws an error if the custom property does not exist.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="key">The key that identifies the custom property object.</param>
        abstract getItem: key: string -> Word.CustomProperty
        /// <summary>Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="key">Required. The key that identifies the custom property object.</param>
        abstract getItemOrNullObject: key: string -> Word.CustomProperty
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.CustomPropertyCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.CustomPropertyCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.CustomPropertyCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.CustomPropertyCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.CustomPropertyCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.CustomPropertyCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.CustomPropertyCollectionData

    /// Contains the collection of {@link Word.CustomProperty} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] CustomPropertyCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> CustomPropertyCollection

    /// The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Document =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract body: Word.Body
        /// Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the properties of the document. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract properties: Word.DocumentProperties
        /// Gets the collection of section objects in the document. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract sections: Word.SectionCollection
        /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract saved: bool
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.DocumentUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Document -> unit
        /// Gets the current selection of the document. Multiple selections are not supported.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getSelection: unit -> Word.Range
        /// Saves the document. This uses the Word default file naming convention if the document has not been saved before.
        /// 
        /// [Api set: WordApi 1.1]
        abstract save: unit -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.DocumentLoadOptions -> Word.Document
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Document
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: DocumentLoadPropertyNamesAndPaths -> Word.Document
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Document
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Document
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Document object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.DocumentData

    type [<AllowNullLiteral>] DocumentLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] DocumentStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Document

    /// The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] DocumentCreated =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract body: Word.Body
        /// Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the properties of the document. Read-only.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract properties: Word.DocumentProperties
        /// Gets the collection of section objects in the document. Read-only.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract sections: Word.SectionCollection
        /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract saved: bool
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.DocumentCreatedUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.DocumentCreated -> unit
        /// Opens the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract ``open``: unit -> unit
        /// Saves the document. This uses the Word default file naming convention if the document has not been saved before.
        /// 
        /// [Api set: WordApiHiddenDocument 1.3]
        abstract save: unit -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.DocumentCreatedLoadOptions -> Word.DocumentCreated
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.DocumentCreated
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: DocumentCreatedLoadPropertyNamesAndPaths -> Word.DocumentCreated
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.DocumentCreated
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.DocumentCreated
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.DocumentCreated object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentCreatedData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.DocumentCreatedData

    type [<AllowNullLiteral>] DocumentCreatedLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] DocumentCreatedStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> DocumentCreated

    /// Represents document properties.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] DocumentProperties =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of custom properties of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract customProperties: Word.CustomPropertyCollection
        /// Gets the application name of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract applicationName: string
        /// Gets or sets the author of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract author: string with get, set
        /// Gets or sets the category of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract category: string with get, set
        /// Gets or sets the comments of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract comments: string with get, set
        /// Gets or sets the company of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract company: string with get, set
        /// Gets the creation date of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract creationDate: DateTime
        /// Gets or sets the format of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract format: string with get, set
        /// Gets or sets the keywords of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract keywords: string with get, set
        /// Gets the last author of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lastAuthor: string
        /// Gets the last print date of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lastPrintDate: DateTime
        /// Gets the last save time of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lastSaveTime: DateTime
        /// Gets or sets the manager of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract manager: string with get, set
        /// Gets the revision number of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract revisionNumber: string
        /// Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
        /// 
        /// [Api set: WordApi 1.3]
        abstract security: float
        /// Gets or sets the subject of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract subject: string with get, set
        /// Gets the template of the document. Read only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract template: string
        /// Gets or sets the title of the document.
        /// 
        /// [Api set: WordApi 1.3]
        abstract title: string with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.DocumentPropertiesUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.DocumentProperties -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.DocumentPropertiesLoadOptions -> Word.DocumentProperties
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.DocumentProperties
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: DocumentPropertiesLoadPropertyNamesAndPaths -> Word.DocumentProperties
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.DocumentProperties
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.DocumentProperties
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.DocumentProperties object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentPropertiesData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.DocumentPropertiesData

    type [<AllowNullLiteral>] DocumentPropertiesLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents document properties.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] DocumentPropertiesStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> DocumentProperties

    /// Represents a font.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Font =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract bold: bool with get, set
        /// Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
        /// 
        /// [Api set: WordApi 1.1]
        abstract color: string with get, set
        /// Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract doubleStrikeThrough: bool with get, set
        /// Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
        ///           *Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
        /// 
        /// [Api set: WordApi 1.1]
        abstract highlightColor: string with get, set
        /// Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract italic: bool with get, set
        /// Gets or sets a value that represents the name of the font.
        /// 
        /// [Api set: WordApi 1.1]
        abstract name: string with get, set
        /// Gets or sets a value that represents the font size in points.
        /// 
        /// [Api set: WordApi 1.1]
        abstract size: float with get, set
        /// Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract strikeThrough: bool with get, set
        /// Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract subscript: bool with get, set
        /// Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
        /// 
        /// [Api set: WordApi 1.1]
        abstract superscript: bool with get, set
        /// Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
        /// 
        /// [Api set: WordApi 1.1]
        abstract underline: U2<Word.UnderlineType, string> with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.FontUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Font -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.FontLoadOptions -> Word.Font
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Font
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: FontLoadPropertyNamesAndPaths -> Word.Font
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Font
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Font
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Font object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FontData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.FontData

    type [<AllowNullLiteral>] FontLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a font.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] FontStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Font

    /// Represents an inline picture.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] InlinePicture =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the parent paragraph that contains the inline image. Read-only.
        /// 
        /// [Api set: WordApi 1.2]
        abstract paragraph: Word.Paragraph
        /// Gets the content control that contains the inline image. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the table that contains the inline image. Throws an error if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCell: Word.TableCell
        /// Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCellOrNullObject: Word.TableCell
        /// Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableOrNullObject: Word.Table
        /// Gets or sets a string that represents the alternative text associated with the inline image.
        /// 
        /// [Api set: WordApi 1.1]
        abstract altTextDescription: string with get, set
        /// Gets or sets a string that contains the title for the inline image.
        /// 
        /// [Api set: WordApi 1.1]
        abstract altTextTitle: string with get, set
        /// Gets or sets a number that describes the height of the inline image.
        /// 
        /// [Api set: WordApi 1.1]
        abstract height: float with get, set
        /// Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
        /// 
        /// [Api set: WordApi 1.1]
        abstract hyperlink: string with get, set
        /// Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
        /// 
        /// [Api set: WordApi 1.1]
        abstract lockAspectRatio: bool with get, set
        /// Gets or sets a number that describes the width of the inline image.
        /// 
        /// [Api set: WordApi 1.1]
        abstract width: float with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.InlinePictureUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.InlinePicture -> unit
        /// Deletes the inline picture from the document.
        /// 
        /// [Api set: WordApi 1.2]
        abstract delete: unit -> unit
        /// Gets the base64 encoded string representation of the inline image.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getBase64ImageSrc: unit -> OfficeExtension.ClientResult<string>
        /// Gets the next inline image. Throws an error if this inline image is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.InlinePicture
        /// Gets the next inline image. Returns a null object if this inline image is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.InlinePicture
        /// <summary>Gets the picture, or the starting or ending point of the picture, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', or 'End'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Gets the picture, or the starting or ending point of the picture, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', or 'End'.</param>
        abstract getRange: ?rangeLocation: InlinePictureGetRangeRangeLocation -> Word.Range
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="breakType">Required. The break type to add.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: Word.BreakType * insertLocation: Word.InsertLocation -> unit
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="breakType">Required. The break type to add.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: InlinePictureInsertBreakBreakType * insertLocation: InlinePictureInsertBreakInsertLocation -> unit
        /// Wraps the inline picture with a rich text content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract insertContentControl: unit -> Word.ContentControl
        /// <summary>Inserts a document at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts a document at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: InlinePictureInsertFileFromBase64InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="html">Required. The HTML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertHtml: html: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="html">Required. The HTML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertHtml: html: string * insertLocation: InlinePictureInsertHtmlInsertLocation -> Word.Range
        /// <summary>Inserts an inline picture at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Before', or 'After'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: Word.InsertLocation -> Word.InlinePicture
        /// <summary>Inserts an inline picture at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Before', or 'After'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: InlinePictureInsertInlinePictureFromBase64InsertLocation -> Word.InlinePicture
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: InlinePictureInsertOoxmlInsertLocation -> Word.Range
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: InlinePictureInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts text at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertText: text: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts text at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertText: text: string * insertLocation: InlinePictureInsertTextInsertLocation -> Word.Range
        /// <summary>Selects the inline picture. This causes Word to scroll to the selection.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects the inline picture. This causes Word to scroll to the selection.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: InlinePictureSelectSelectionMode -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.InlinePictureLoadOptions -> Word.InlinePicture
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.InlinePicture
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: InlinePictureLoadPropertyNamesAndPaths -> Word.InlinePicture
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.InlinePicture
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.InlinePicture
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.InlinePicture object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.InlinePictureData

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertBreakBreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertBreakInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertFileFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertHtmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertInlinePictureFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertOoxmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureInsertTextInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] InlinePictureSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<AllowNullLiteral>] InlinePictureLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents an inline picture.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] InlinePictureStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InlinePicture

    /// Contains a collection of {@link Word.InlinePicture} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] InlinePictureCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.InlinePicture>
        /// Gets the first inline image in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.InlinePicture
        /// Gets the first inline image in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.InlinePicture
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.InlinePictureCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.InlinePictureCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.InlinePictureCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.InlinePictureCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.InlinePictureCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.InlinePictureCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.InlinePictureCollectionData

    /// Contains a collection of {@link Word.InlinePicture} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] InlinePictureCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> InlinePictureCollection

    /// Contains a collection of {@link Word.Paragraph} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] List =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets paragraphs in the list. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract paragraphs: Word.ParagraphCollection
        /// Gets the list's id.
        /// 
        /// [Api set: WordApi 1.3]
        abstract id: float
        /// Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract levelExistences: ResizeArray<bool>
        /// Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract levelTypes: ResizeArray<Word.ListLevelType>
        /// <summary>Gets the paragraphs that occur at the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        abstract getLevelParagraphs: level: float -> Word.ParagraphCollection
        /// <summary>Gets the bullet, number, or picture at the specified level as a string.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        abstract getLevelString: level: float -> OfficeExtension.ClientResult<string>
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: ListInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Sets the alignment of the bullet, number, or picture at the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="alignment">Required. The level alignment that can be 'Left', 'Centered', or 'Right'.</param>
        abstract setLevelAlignment: level: float * alignment: Word.Alignment -> unit
        /// <summary>Sets the alignment of the bullet, number, or picture at the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="alignment">Required. The level alignment that can be 'Left', 'Centered', or 'Right'.</param>
        abstract setLevelAlignment: level: float * alignment: ListSetLevelAlignmentAlignment -> unit
        /// <summary>Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="listBullet">Required. The bullet.</param>
        /// <param name="charCode">Optional. The bullet character's code value. Used only if the bullet is 'Custom'.</param>
        /// <param name="fontName">Optional. The bullet's font name. Used only if the bullet is 'Custom'.</param>
        abstract setLevelBullet: level: float * listBullet: Word.ListBullet * ?charCode: float * ?fontName: string -> unit
        /// <summary>Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="listBullet">Required. The bullet.</param>
        /// <param name="charCode">Optional. The bullet character's code value. Used only if the bullet is 'Custom'.</param>
        /// <param name="fontName">Optional. The bullet's font name. Used only if the bullet is 'Custom'.</param>
        abstract setLevelBullet: level: float * listBullet: ListSetLevelBulletListBullet * ?charCode: float * ?fontName: string -> unit
        /// <summary>Sets the two indents of the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="textIndent">Required. The text indent in points. It is the same as paragraph left indent.</param>
        /// <param name="bulletNumberPictureIndent">Required. The relative indent, in points, of the bullet, number, or picture. It is the same as paragraph first line indent.</param>
        abstract setLevelIndents: level: float * textIndent: float * bulletNumberPictureIndent: float -> unit
        /// <summary>Sets the numbering format at the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="listNumbering">Required. The ordinal format.</param>
        /// <param name="formatString">Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.</param>
        abstract setLevelNumbering: level: float * listNumbering: Word.ListNumbering * ?formatString: Array<U2<string, float>> -> unit
        /// <summary>Sets the numbering format at the specified level in the list.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="listNumbering">Required. The ordinal format.</param>
        /// <param name="formatString">Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.</param>
        abstract setLevelNumbering: level: float * listNumbering: ListSetLevelNumberingListNumbering * ?formatString: Array<U2<string, float>> -> unit
        /// <summary>Sets the starting number at the specified level in the list. Default value is 1.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="level">Required. The level in the list.</param>
        /// <param name="startingNumber">Required. The number to start with.</param>
        abstract setLevelStartingNumber: level: float * startingNumber: float -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.ListLoadOptions -> Word.List
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.List
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: ListLoadPropertyNamesAndPaths -> Word.List
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.List
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.List
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.List object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.ListData

    type [<StringEnum>] [<RequireQualifiedAccess>] ListInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ListSetLevelAlignmentAlignment =
        | [<CompiledName "Mixed">] Mixed
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "Left">] Left
        | [<CompiledName "Centered">] Centered
        | [<CompiledName "Right">] Right
        | [<CompiledName "Justified">] Justified

    type [<StringEnum>] [<RequireQualifiedAccess>] ListSetLevelBulletListBullet =
        | [<CompiledName "Custom">] Custom
        | [<CompiledName "Solid">] Solid
        | [<CompiledName "Hollow">] Hollow
        | [<CompiledName "Square">] Square
        | [<CompiledName "Diamonds">] Diamonds
        | [<CompiledName "Arrow">] Arrow
        | [<CompiledName "Checkmark">] Checkmark

    type [<StringEnum>] [<RequireQualifiedAccess>] ListSetLevelNumberingListNumbering =
        | [<CompiledName "None">] None
        | [<CompiledName "Arabic">] Arabic
        | [<CompiledName "UpperRoman">] UpperRoman
        | [<CompiledName "LowerRoman">] LowerRoman
        | [<CompiledName "UpperLetter">] UpperLetter
        | [<CompiledName "LowerLetter">] LowerLetter

    type [<AllowNullLiteral>] ListLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Contains a collection of {@link Word.Paragraph} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ListStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> List

    /// Contains a collection of {@link Word.List} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ListCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.List>
        /// <summary>Gets a list by its identifier. Throws an error if there isn't a list with the identifier in this collection.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="id">Required. A list identifier.</param>
        abstract getById: id: float -> Word.List
        /// <summary>Gets a list by its identifier. Returns a null object if there isn't a list with the identifier in this collection.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="id">Required. A list identifier.</param>
        abstract getByIdOrNullObject: id: float -> Word.List
        /// Gets the first list in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.List
        /// Gets the first list in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.List
        /// <summary>Gets a list object by its index in the collection.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="index">A number that identifies the index location of a list object.</param>
        abstract getItem: index: float -> Word.List
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.ListCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.ListCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.ListCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.ListCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.ListCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.ListCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.ListCollectionData

    /// Contains a collection of {@link Word.List} objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ListCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ListCollection

    /// Represents the paragraph list item format.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ListItem =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets or sets the level of the item in the list.
        /// 
        /// [Api set: WordApi 1.3]
        abstract level: float with get, set
        /// Gets the list item bullet, number, or picture as a string. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract listString: string
        /// Gets the list item order number in relation to its siblings. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract siblingIndex: float
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.ListItemUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.ListItem -> unit
        /// <summary>Gets the list item parent, or the closest ancestor if the parent does not exist. Throws an error if the list item has no ancestor.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="parentOnly">Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.</param>
        abstract getAncestor: ?parentOnly: bool -> Word.Paragraph
        /// <summary>Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancestor.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="parentOnly">Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.</param>
        abstract getAncestorOrNullObject: ?parentOnly: bool -> Word.Paragraph
        /// <summary>Gets all descendant list items of the list item.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="directChildrenOnly">Optional. Specifies only the list item's direct children will be returned. The default is false that indicates to get all descendant items.</param>
        abstract getDescendants: ?directChildrenOnly: bool -> Word.ParagraphCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.ListItemLoadOptions -> Word.ListItem
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.ListItem
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: ListItemLoadPropertyNamesAndPaths -> Word.ListItem
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.ListItem
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.ListItem
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.ListItem object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListItemData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.ListItemData

    type [<AllowNullLiteral>] ListItemLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents the paragraph list item format.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] ListItemStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ListItem

    /// Represents a single paragraph in a selection, range, content control, or document body.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Paragraph =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of content control objects in the paragraph. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract font: Word.Font
        /// Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract inlinePictures: Word.InlinePictureCollection
        /// Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract list: Word.List
        /// Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract listItem: Word.ListItem
        /// Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract listItemOrNullObject: Word.ListItem
        /// Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract listOrNullObject: Word.List
        /// Gets the parent body of the paragraph. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBody: Word.Body
        /// Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the table that contains the paragraph. Throws an error if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCell: Word.TableCell
        /// Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCellOrNullObject: Word.TableCell
        /// Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableOrNullObject: Word.Table
        /// Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
        /// 
        /// [Api set: WordApi 1.1]
        abstract alignment: U2<Word.Alignment, string> with get, set
        /// Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
        /// 
        /// [Api set: WordApi 1.1]
        abstract firstLineIndent: float with get, set
        /// Indicates the paragraph is the last one inside its parent body. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract isLastParagraph: bool
        /// Checks whether the paragraph is a list item. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract isListItem: bool
        /// Gets or sets the left indent value, in points, for the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract leftIndent: float with get, set
        /// Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
        /// 
        /// [Api set: WordApi 1.1]
        abstract lineSpacing: float with get, set
        /// Gets or sets the amount of spacing, in grid lines, after the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract lineUnitAfter: float with get, set
        /// Gets or sets the amount of spacing, in grid lines, before the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract lineUnitBefore: float with get, set
        /// Gets or sets the outline level for the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract outlineLevel: float with get, set
        /// Gets or sets the right indent value, in points, for the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract rightIndent: float with get, set
        /// Gets or sets the spacing, in points, after the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract spaceAfter: float with get, set
        /// Gets or sets the spacing, in points, before the paragraph.
        /// 
        /// [Api set: WordApi 1.1]
        abstract spaceBefore: float with get, set
        /// Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
        /// 
        /// [Api set: WordApi 1.1]
        abstract style: string with get, set
        /// Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBuiltIn: U2<Word.Style, string> with get, set
        /// Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract tableNestingLevel: float
        /// Gets the text of the paragraph. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract text: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.ParagraphUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Paragraph -> unit
        /// <summary>Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="listId">Required. The ID of an existing list.</param>
        /// <param name="level">Required. The level in the list.</param>
        abstract attachToList: listId: float * level: float -> Word.List
        /// Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
        /// 
        /// [Api set: WordApi 1.1]
        abstract clear: unit -> unit
        /// Deletes the paragraph and its content from the document.
        /// 
        /// [Api set: WordApi 1.1]
        abstract delete: unit -> unit
        /// Moves this paragraph out of its list, if the paragraph is a list item.
        /// 
        /// [Api set: WordApi 1.3]
        abstract detachFromList: unit -> unit
        /// Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getHtml: unit -> OfficeExtension.ClientResult<string>
        /// Gets the next paragraph. Throws an error if the paragraph is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.Paragraph
        /// Gets the next paragraph. Returns a null object if the paragraph is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.Paragraph
        /// Gets the Office Open XML (OOXML) representation of the paragraph object.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getOoxml: unit -> OfficeExtension.ClientResult<string>
        /// Gets the previous paragraph. Throws an error if the paragraph is the first one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getPrevious: unit -> Word.Paragraph
        /// Gets the previous paragraph. Returns a null object if the paragraph is the first one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getPreviousOrNullObject: unit -> Word.Paragraph
        /// <summary>Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: ParagraphGetRangeRangeLocation -> Word.Range
        /// <summary>Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract getTextRanges: endingMarks: ResizeArray<string> * ?trimSpacing: bool -> Word.RangeCollection
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add to the document.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: Word.BreakType * insertLocation: Word.InsertLocation -> unit
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add to the document.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: ParagraphInsertBreakBreakType * insertLocation: ParagraphInsertBreakInsertLocation -> unit
        /// Wraps the paragraph object with a rich text content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract insertContentControl: unit -> Word.ContentControl
        /// <summary>Inserts a document into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts a document into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: ParagraphInsertFileFromBase64InsertLocation -> Word.Range
        /// <summary>Inserts HTML into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in the paragraph.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertHtml: html: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts HTML into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted in the paragraph.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertHtml: html: string * insertLocation: ParagraphInsertHtmlInsertLocation -> Word.Range
        /// <summary>Inserts a picture into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: Word.InsertLocation -> Word.InlinePicture
        /// <summary>Inserts a picture into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: ParagraphInsertInlinePictureFromBase64InsertLocation -> Word.InlinePicture
        /// <summary>Inserts OOXML into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted in the paragraph.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts OOXML into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted in the paragraph.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: ParagraphInsertOoxmlInsertLocation -> Word.Range
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: ParagraphInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: Word.InsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: ParagraphInsertTableInsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts text into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertText: text: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts text into the paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', or 'End'.</param>
        abstract insertText: text: string * insertLocation: ParagraphInsertTextInsertLocation -> Word.Range
        /// <summary>Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="searchText">Required. The search text.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects and navigates the Word UI to the paragraph.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects and navigates the Word UI to the paragraph.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: ParagraphSelectSelectionMode -> unit
        /// <summary>Splits the paragraph into child ranges by using delimiters.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="delimiters">Required. The delimiters as an array of strings.</param>
        /// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract split: delimiters: ResizeArray<string> * ?trimDelimiters: bool * ?trimSpacing: bool -> Word.RangeCollection
        /// Starts a new list with this paragraph. Fails if the paragraph is already a list item.
        /// 
        /// [Api set: WordApi 1.3]
        abstract startNewList: unit -> Word.List
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.ParagraphLoadOptions -> Word.Paragraph
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Paragraph
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: ParagraphLoadPropertyNamesAndPaths -> Word.Paragraph
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Paragraph
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Paragraph
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Paragraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.ParagraphData

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertBreakBreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertBreakInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertFileFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertHtmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertInlinePictureFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertOoxmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertTableInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphInsertTextInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] ParagraphSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<AllowNullLiteral>] ParagraphLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a single paragraph in a selection, range, content control, or document body.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ParagraphStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Paragraph

    /// Contains a collection of {@link Word.Paragraph} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ParagraphCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.Paragraph>
        /// Gets the first paragraph in this collection. Throws an error if the collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.Paragraph
        /// Gets the first paragraph in this collection. Returns a null object if the collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.Paragraph
        /// Gets the last paragraph in this collection. Throws an error if the collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getLast: unit -> Word.Paragraph
        /// Gets the last paragraph in this collection. Returns a null object if the collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getLastOrNullObject: unit -> Word.Paragraph
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.ParagraphCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.ParagraphCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.ParagraphCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.ParagraphCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.ParagraphCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.ParagraphCollectionData

    /// Contains a collection of {@link Word.Paragraph} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] ParagraphCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> ParagraphCollection

    /// Represents a contiguous area in a document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Range =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the collection of content control objects in the range. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract contentControls: Word.ContentControlCollection
        /// Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract font: Word.Font
        /// Gets the collection of inline picture objects in the range. Read-only.
        /// 
        /// [Api set: WordApi 1.2]
        abstract inlinePictures: Word.InlinePictureCollection
        /// Gets the collection of list objects in the range. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract lists: Word.ListCollection
        /// Gets the collection of paragraph objects in the range. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract paragraphs: Word.ParagraphCollection
        /// Gets the parent body of the range. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBody: Word.Body
        /// Gets the content control that contains the range. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the table that contains the range. Throws an error if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the table cell that contains the range. Throws an error if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCell: Word.TableCell
        /// Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCellOrNullObject: Word.TableCell
        /// Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableOrNullObject: Word.Table
        /// Gets the collection of table objects in the range. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract tables: Word.TableCollection
        /// Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
        /// 
        /// [Api set: WordApi 1.3]
        abstract hyperlink: string with get, set
        /// Checks whether the range length is zero. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract isEmpty: bool
        /// Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
        /// 
        /// [Api set: WordApi 1.1]
        abstract style: string with get, set
        /// Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBuiltIn: U2<Word.Style, string> with get, set
        /// Gets the text of the range. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract text: string
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.RangeUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Range -> unit
        /// Clears the contents of the range object. The user can perform the undo operation on the cleared content.
        /// 
        /// [Api set: WordApi 1.1]
        abstract clear: unit -> unit
        /// <summary>Compares this range's location with another range's location.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="range">Required. The range to compare with this range.</param>
        abstract compareLocationWith: range: Word.Range -> OfficeExtension.ClientResult<Word.LocationRelation>
        /// Deletes the range and its content from the document.
        /// 
        /// [Api set: WordApi 1.1]
        abstract delete: unit -> unit
        /// <summary>Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Throws an error if the two ranges do not have a union.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="range">Required. Another range.</param>
        abstract expandTo: range: Word.Range -> Word.Range
        /// <summary>Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Returns a null object if the two ranges do not have a union.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="range">Required. Another range.</param>
        abstract expandToOrNullObject: range: Word.Range -> Word.Range
        /// Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getHtml: unit -> OfficeExtension.ClientResult<string>
        /// Gets hyperlink child ranges within the range.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getHyperlinkRanges: unit -> Word.RangeCollection
        /// <summary>Gets the next text range by using punctuation marks and/or other ending marks. Throws an error if this text range is the last one.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.</param>
        abstract getNextTextRange: endingMarks: ResizeArray<string> * ?trimSpacing: bool -> Word.Range
        /// <summary>Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.</param>
        abstract getNextTextRangeOrNullObject: endingMarks: ResizeArray<string> * ?trimSpacing: bool -> Word.Range
        /// Gets the OOXML representation of the range object.
        /// 
        /// [Api set: WordApi 1.1]
        abstract getOoxml: unit -> OfficeExtension.ClientResult<string>
        /// <summary>Clones the range, or gets the starting or ending point of the range as a new range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Clones the range, or gets the starting or ending point of the range as a new range.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.</param>
        abstract getRange: ?rangeLocation: RangeGetRangeRangeLocation -> Word.Range
        /// <summary>Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract getTextRanges: endingMarks: ResizeArray<string> * ?trimSpacing: bool -> Word.RangeCollection
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: Word.BreakType * insertLocation: Word.InsertLocation -> unit
        /// <summary>Inserts a break at the specified location in the main document.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="breakType">Required. The break type to add.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertBreak: breakType: RangeInsertBreakBreakType * insertLocation: RangeInsertBreakInsertLocation -> unit
        /// Wraps the range object with a rich text content control.
        /// 
        /// [Api set: WordApi 1.1]
        abstract insertContentControl: unit -> Word.ContentControl
        /// <summary>Inserts a document at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts a document at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertFileFromBase64: base64File: string * insertLocation: RangeInsertFileFromBase64InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertHtml: html: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts HTML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="html">Required. The HTML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertHtml: html: string * insertLocation: RangeInsertHtmlInsertLocation -> Word.Range
        /// <summary>Inserts a picture at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: Word.InsertLocation -> Word.InlinePicture
        /// <summary>Inserts a picture at the specified location.
        /// 
        /// [Api set: WordApi 1.2]</summary>
        /// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertInlinePictureFromBase64: base64EncodedImage: string * insertLocation: RangeInsertInlinePictureFromBase64InsertLocation -> Word.InlinePicture
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts OOXML at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="ooxml">Required. The OOXML to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertOoxml: ooxml: string * insertLocation: RangeInsertOoxmlInsertLocation -> Word.Range
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: RangeInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: Word.InsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: RangeInsertTableInsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts text at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertText: text: string * insertLocation: Word.InsertLocation -> Word.Range
        /// <summary>Inserts text at the specified location.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="text">Required. Text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.</param>
        abstract insertText: text: string * insertLocation: RangeInsertTextInsertLocation -> Word.Range
        /// <summary>Returns a new range as the intersection of this range with another range. This range is not changed. Throws an error if the two ranges are not overlapped or adjacent.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="range">Required. Another range.</param>
        abstract intersectWith: range: Word.Range -> Word.Range
        /// <summary>Returns a new range as the intersection of this range with another range. This range is not changed. Returns a null object if the two ranges are not overlapped or adjacent.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="range">Required. Another range.</param>
        abstract intersectWithOrNullObject: range: Word.Range -> Word.Range
        /// <summary>Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="searchText">Required. The search text.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects and navigates the Word UI to the range.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects and navigates the Word UI to the range.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: RangeSelectSelectionMode -> unit
        /// <summary>Splits the range into child ranges by using delimiters.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="delimiters">Required. The delimiters as an array of strings.</param>
        /// <param name="multiParagraphs">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
        /// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
        /// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
        abstract split: delimiters: ResizeArray<string> * ?multiParagraphs: bool * ?trimDelimiters: bool * ?trimSpacing: bool -> Word.RangeCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.RangeLoadOptions -> Word.Range
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Range
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: RangeLoadPropertyNamesAndPaths -> Word.Range
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Range
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Range
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Range object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.RangeData

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertBreakBreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertBreakInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertFileFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertHtmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertInlinePictureFromBase64InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertOoxmlInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertTableInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeInsertTextInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<AllowNullLiteral>] RangeLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a contiguous area in a document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] RangeStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Range

    /// Contains a collection of {@link Word.Range} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] RangeCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.Range>
        /// Gets the first range in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.Range
        /// Gets the first range in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.Range
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.RangeCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.RangeCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.RangeCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.RangeCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.RangeCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.RangeCollectionData

    /// Contains a collection of {@link Word.Range} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] RangeCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> RangeCollection

    /// Specifies the options to be included in a search operation.
    /// 
    /// To learn more about how to use search options in the Word JavaScript APIs, read {@link https://docs.microsoft.com/office/dev/add-ins/word/search-option-guidance | Use search options to find text in your Word add-in}.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] SearchOptions =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract ignorePunct: bool with get, set
        /// Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract ignoreSpace: bool with get, set
        /// Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract matchCase: bool with get, set
        /// Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract matchPrefix: bool with get, set
        /// Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract matchSuffix: bool with get, set
        /// Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract matchWholeWord: bool with get, set
        /// Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
        /// 
        /// [Api set: WordApi 1.1]
        abstract matchWildcards: bool with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.SearchOptionsUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.SearchOptions -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.SearchOptionsLoadOptions -> Word.SearchOptions
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.SearchOptions
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: SearchOptionsLoadPropertyNamesAndPaths -> Word.SearchOptions
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.SearchOptions object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.SearchOptionsData

    type [<AllowNullLiteral>] SearchOptionsLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Specifies the options to be included in a search operation.
    /// 
    /// To learn more about how to use search options in the Word JavaScript APIs, read {@link https://docs.microsoft.com/office/dev/add-ins/word/search-option-guidance | Use search options to find text in your Word add-in}.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] SearchOptionsStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SearchOptions
        /// Create a new instance of Word.SearchOptions object
        abstract newObject: context: OfficeExtension.ClientRequestContext -> Word.SearchOptions

    /// Represents a section in a Word document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] Section =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
        /// 
        /// [Api set: WordApi 1.1]
        abstract body: Word.Body
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.SectionUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Section -> unit
        /// <summary>Gets one of the section's footers.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="type">Required. The type of footer to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.</param>
        abstract getFooter: ``type``: Word.HeaderFooterType -> Word.Body
        /// <summary>Gets one of the section's footers.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="type">Required. The type of footer to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.</param>
        abstract getFooter: ``type``: SectionGetFooterType -> Word.Body
        /// <summary>Gets one of the section's headers.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="type">Required. The type of header to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.</param>
        abstract getHeader: ``type``: Word.HeaderFooterType -> Word.Body
        /// <summary>Gets one of the section's headers.
        /// 
        /// [Api set: WordApi 1.1]</summary>
        /// <param name="type">Required. The type of header to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.</param>
        abstract getHeader: ``type``: SectionGetHeaderType -> Word.Body
        /// Gets the next section. Throws an error if this section is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.Section
        /// Gets the next section. Returns a null object if this section is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.Section
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.SectionLoadOptions -> Word.Section
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Section
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: SectionLoadPropertyNamesAndPaths -> Word.Section
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Section
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Section
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Section object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.SectionData

    type [<StringEnum>] [<RequireQualifiedAccess>] SectionGetFooterType =
        | [<CompiledName "Primary">] Primary
        | [<CompiledName "FirstPage">] FirstPage
        | [<CompiledName "EvenPages">] EvenPages

    type [<StringEnum>] [<RequireQualifiedAccess>] SectionGetHeaderType =
        | [<CompiledName "Primary">] Primary
        | [<CompiledName "FirstPage">] FirstPage
        | [<CompiledName "EvenPages">] EvenPages

    type [<AllowNullLiteral>] SectionLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a section in a Word document.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] SectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Section

    /// Contains the collection of the document's {@link Word.Section} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] SectionCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.Section>
        /// Gets the first section in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.Section
        /// Gets the first section in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.Section
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.SectionCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.SectionCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.SectionCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.SectionCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.SectionCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.SectionCollectionData

    /// Contains the collection of the document's {@link Word.Section} objects.
    /// 
    /// [Api set: WordApi 1.1]
    type [<AllowNullLiteral>] SectionCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> SectionCollection

    /// Represents a table in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] Table =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract font: Word.Font
        /// Gets the parent body of the table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentBody: Word.Body
        /// Gets the content control that contains the table. Throws an error if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControl: Word.ContentControl
        /// Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentContentControlOrNullObject: Word.ContentControl
        /// Gets the table that contains this table. Throws an error if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the table cell that contains this table. Throws an error if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCell: Word.TableCell
        /// Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableCellOrNullObject: Word.TableCell
        /// Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTableOrNullObject: Word.Table
        /// Gets all of the table rows. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract rows: Word.TableRowCollection
        /// Gets the child tables nested one level deeper. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract tables: Word.TableCollection
        /// Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract alignment: U2<Word.Alignment, string> with get, set
        /// Gets and sets the number of header rows.
        /// 
        /// [Api set: WordApi 1.3]
        abstract headerRowCount: float with get, set
        /// Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract horizontalAlignment: U2<Word.Alignment, string> with get, set
        /// Indicates whether all of the table rows are uniform. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract isUniform: bool
        /// Gets the nesting level of the table. Top-level tables have level 1. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract nestingLevel: float
        /// Gets the number of rows in the table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract rowCount: float
        /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
        /// 
        /// [Api set: WordApi 1.3]
        abstract shadingColor: string with get, set
        /// Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract style: string with get, set
        /// Gets and sets whether the table has banded columns.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBandedColumns: bool with get, set
        /// Gets and sets whether the table has banded rows.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBandedRows: bool with get, set
        /// Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleBuiltIn: U2<Word.Style, string> with get, set
        /// Gets and sets whether the table has a first column with a special style.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleFirstColumn: bool with get, set
        /// Gets and sets whether the table has a last column with a special style.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleLastColumn: bool with get, set
        /// Gets and sets whether the table has a total (last) row with a special style.
        /// 
        /// [Api set: WordApi 1.3]
        abstract styleTotalRow: bool with get, set
        /// Gets and sets the text values in the table, as a 2D Javascript array.
        /// 
        /// [Api set: WordApi 1.3]
        abstract values: ResizeArray<ResizeArray<string>> with get, set
        /// Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract verticalAlignment: U2<Word.VerticalAlignment, string> with get, set
        /// Gets and sets the width of the table in points.
        /// 
        /// [Api set: WordApi 1.3]
        abstract width: float with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.Table -> unit
        /// <summary>Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table.</param>
        /// <param name="columnCount">Required. Number of columns to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract addColumns: insertLocation: Word.InsertLocation * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> unit
        /// <summary>Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table.</param>
        /// <param name="columnCount">Required. Number of columns to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract addColumns: insertLocation: TableAddColumnsInsertLocation * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> unit
        /// <summary>Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Start' or 'End'.</param>
        /// <param name="rowCount">Required. Number of rows to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract addRows: insertLocation: Word.InsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Start' or 'End'.</param>
        /// <param name="rowCount">Required. Number of rows to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract addRows: insertLocation: TableAddRowsInsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// Autofits the table columns to the width of the window.
        /// 
        /// [Api set: WordApi 1.3]
        abstract autoFitWindow: unit -> unit
        /// Clears the contents of the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract clear: unit -> unit
        /// Deletes the entire table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract delete: unit -> unit
        /// <summary>Deletes specific columns. This is applicable to uniform tables.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="columnIndex">Required. The first column to delete.</param>
        /// <param name="columnCount">Optional. The number of columns to delete. Default 1.</param>
        abstract deleteColumns: columnIndex: float * ?columnCount: float -> unit
        /// <summary>Deletes specific rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowIndex">Required. The first row to delete.</param>
        /// <param name="rowCount">Optional. The number of rows to delete. Default 1.</param>
        abstract deleteRows: rowIndex: float * ?rowCount: float -> unit
        /// Distributes the column widths evenly. This is applicable to uniform tables.
        /// 
        /// [Api set: WordApi 1.3]
        abstract distributeColumns: unit -> unit
        /// <summary>Gets the border style for the specified border.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: Word.BorderLocation -> Word.TableBorder
        /// <summary>Gets the border style for the specified border.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: TableGetBorderBorderLocation -> Word.TableBorder
        /// <summary>Gets the table cell at a specified row and column. Throws an error if the specified table cell does not exist.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowIndex">Required. The index of the row.</param>
        /// <param name="cellIndex">Required. The index of the cell in the row.</param>
        abstract getCell: rowIndex: float * cellIndex: float -> Word.TableCell
        /// <summary>Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowIndex">Required. The index of the row.</param>
        /// <param name="cellIndex">Required. The index of the cell in the row.</param>
        abstract getCellOrNullObject: rowIndex: float * cellIndex: float -> Word.TableCell
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: Word.CellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: TableGetCellPaddingCellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// Gets the next table. Throws an error if this table is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.Table
        /// Gets the next table. Returns a null object if this table is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.Table
        /// Gets the paragraph after the table. Throws an error if there isn't a paragraph after the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getParagraphAfter: unit -> Word.Paragraph
        /// Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getParagraphAfterOrNullObject: unit -> Word.Paragraph
        /// Gets the paragraph before the table. Throws an error if there isn't a paragraph before the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getParagraphBefore: unit -> Word.Paragraph
        /// Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getParagraphBeforeOrNullObject: unit -> Word.Paragraph
        /// <summary>Gets the range that contains this table, or the range at the start or end of the table.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', or 'After'.</param>
        abstract getRange: ?rangeLocation: Word.RangeLocation -> Word.Range
        /// <summary>Gets the range that contains this table, or the range at the start or end of the table.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', or 'After'.</param>
        abstract getRange: ?rangeLocation: TableGetRangeRangeLocation -> Word.Range
        /// Inserts a content control on the table.
        /// 
        /// [Api set: WordApi 1.3]
        abstract insertContentControl: unit -> Word.ContentControl
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: Word.InsertLocation -> Word.Paragraph
        /// <summary>Inserts a paragraph at the specified location.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        abstract insertParagraph: paragraphText: string * insertLocation: TableInsertParagraphInsertLocation -> Word.Paragraph
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: Word.InsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Inserts a table with the specified number of rows and columns.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="rowCount">Required. The number of rows in the table.</param>
        /// <param name="columnCount">Required. The number of columns in the table.</param>
        /// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertTable: rowCount: float * columnCount: float * insertLocation: TableInsertTableInsertLocation * ?values: ResizeArray<ResizeArray<string>> -> Word.Table
        /// <summary>Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="searchText">Required. The search text.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: TableSelectSelectionMode -> unit
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: Word.CellPaddingLocation * cellPadding: float -> unit
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: TableSetCellPaddingCellPaddingLocation * cellPadding: float -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.TableLoadOptions -> Word.Table
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.Table
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: TableLoadPropertyNamesAndPaths -> Word.Table
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.Table
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.Table
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.Table object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.TableData

    type [<StringEnum>] [<RequireQualifiedAccess>] TableAddColumnsInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableAddRowsInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableGetBorderBorderLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right
        | [<CompiledName "InsideHorizontal">] InsideHorizontal
        | [<CompiledName "InsideVertical">] InsideVertical
        | [<CompiledName "Inside">] Inside
        | [<CompiledName "Outside">] Outside
        | [<CompiledName "All">] All

    type [<StringEnum>] [<RequireQualifiedAccess>] TableGetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<StringEnum>] [<RequireQualifiedAccess>] TableGetRangeRangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] TableInsertParagraphInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableInsertTableInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<StringEnum>] [<RequireQualifiedAccess>] TableSetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<AllowNullLiteral>] TableLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a table in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> Table

    /// Contains the collection of the document's Table objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.Table>
        /// Gets the first table in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.Table
        /// Gets the first table in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.Table
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.TableCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.TableCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.TableCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.TableCollectionData

    /// Contains the collection of the document's Table objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableCollection

    /// Represents a row in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableRow =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets cells. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract cells: Word.TableCellCollection
        /// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract font: Word.Font
        /// Gets parent table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the number of cells in the row. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract cellCount: float
        /// Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract horizontalAlignment: U2<Word.Alignment, string> with get, set
        /// Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
        /// 
        /// [Api set: WordApi 1.3]
        abstract isHeader: bool
        /// Gets and sets the preferred height of the row in points.
        /// 
        /// [Api set: WordApi 1.3]
        abstract preferredHeight: float with get, set
        /// Gets the index of the row in its parent table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract rowIndex: float
        /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
        /// 
        /// [Api set: WordApi 1.3]
        abstract shadingColor: string with get, set
        /// Gets and sets the text values in the row, as a 2D Javascript array.
        /// 
        /// [Api set: WordApi 1.3]
        abstract values: ResizeArray<ResizeArray<string>> with get, set
        /// Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract verticalAlignment: U2<Word.VerticalAlignment, string> with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableRowUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.TableRow -> unit
        /// Clears the contents of the row.
        /// 
        /// [Api set: WordApi 1.3]
        abstract clear: unit -> unit
        /// Deletes the entire row.
        /// 
        /// [Api set: WordApi 1.3]
        abstract delete: unit -> unit
        /// <summary>Gets the border style of the cells in the row.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: Word.BorderLocation -> Word.TableBorder
        /// <summary>Gets the border style of the cells in the row.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: TableRowGetBorderBorderLocation -> Word.TableBorder
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: Word.CellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: TableRowGetCellPaddingCellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// Gets the next row. Throws an error if this row is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.TableRow
        /// Gets the next row. Returns a null object if this row is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.TableRow
        /// <summary>Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'.</param>
        /// <param name="rowCount">Required. Number of rows to add</param>
        /// <param name="values">Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.</param>
        abstract insertRows: insertLocation: Word.InsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'.</param>
        /// <param name="rowCount">Required. Number of rows to add</param>
        /// <param name="values">Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.</param>
        abstract insertRows: insertLocation: TableRowInsertRowsInsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="searchText">Required. The search text.</param>
        /// <param name="searchOptions">Optional. Options for the search.</param>
        abstract search: searchText: string * ?searchOptions: U2<Word.SearchOptions, BodySearch> -> Word.RangeCollection
        /// <summary>Selects the row and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: Word.SelectionMode -> unit
        /// <summary>Selects the row and navigates the Word UI to it.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.</param>
        abstract select: ?selectionMode: TableRowSelectSelectionMode -> unit
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: Word.CellPaddingLocation * cellPadding: float -> unit
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: TableRowSetCellPaddingCellPaddingLocation * cellPadding: float -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.TableRowLoadOptions -> Word.TableRow
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableRow
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: TableRowLoadPropertyNamesAndPaths -> Word.TableRow
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableRow
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableRow
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.TableRow object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.TableRowData

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowGetBorderBorderLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right
        | [<CompiledName "InsideHorizontal">] InsideHorizontal
        | [<CompiledName "InsideVertical">] InsideVertical
        | [<CompiledName "Inside">] Inside
        | [<CompiledName "Outside">] Outside
        | [<CompiledName "All">] All

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowGetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowInsertRowsInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowSelectSelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<StringEnum>] [<RequireQualifiedAccess>] TableRowSetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<AllowNullLiteral>] TableRowLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a row in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableRowStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableRow

    /// Contains the collection of the document's TableRow objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableRowCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.TableRow>
        /// Gets the first row in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.TableRow
        /// Gets the first row in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.TableRow
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.TableRowCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.TableRowCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableRowCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableRowCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.TableRowCollectionData

    /// Contains the collection of the document's TableRow objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableRowCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableRowCollection

    /// Represents a table cell in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCell =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the body object of the cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract body: Word.Body
        /// Gets the parent row of the cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentRow: Word.TableRow
        /// Gets the parent table of the cell. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract parentTable: Word.Table
        /// Gets the index of the cell in its row. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract cellIndex: float
        /// Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
        /// 
        /// [Api set: WordApi 1.3]
        abstract columnWidth: float with get, set
        /// Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract horizontalAlignment: U2<Word.Alignment, string> with get, set
        /// Gets the index of the cell's row in the table. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract rowIndex: float
        /// Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
        /// 
        /// [Api set: WordApi 1.3]
        abstract shadingColor: string with get, set
        /// Gets and sets the text of the cell.
        /// 
        /// [Api set: WordApi 1.3]
        abstract value: string with get, set
        /// Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
        /// 
        /// [Api set: WordApi 1.3]
        abstract verticalAlignment: U2<Word.VerticalAlignment, string> with get, set
        /// Gets the width of the cell in points. Read-only.
        /// 
        /// [Api set: WordApi 1.3]
        abstract width: float
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableCellUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.TableCell -> unit
        /// Deletes the column containing this cell. This is applicable to uniform tables.
        /// 
        /// [Api set: WordApi 1.3]
        abstract deleteColumn: unit -> unit
        /// Deletes the row containing this cell.
        /// 
        /// [Api set: WordApi 1.3]
        abstract deleteRow: unit -> unit
        /// <summary>Gets the border style for the specified border.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: Word.BorderLocation -> Word.TableBorder
        /// <summary>Gets the border style for the specified border.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="borderLocation">Required. The border location.</param>
        abstract getBorder: borderLocation: TableCellGetBorderBorderLocation -> Word.TableBorder
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: Word.CellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// <summary>Gets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        abstract getCellPadding: cellPaddingLocation: TableCellGetCellPaddingCellPaddingLocation -> OfficeExtension.ClientResult<float>
        /// Gets the next cell. Throws an error if this cell is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNext: unit -> Word.TableCell
        /// Gets the next cell. Returns a null object if this cell is the last one.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getNextOrNullObject: unit -> Word.TableCell
        /// <summary>Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
        /// <param name="columnCount">Required. Number of columns to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertColumns: insertLocation: Word.InsertLocation * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> unit
        /// <summary>Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
        /// <param name="columnCount">Required. Number of columns to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertColumns: insertLocation: TableCellInsertColumnsInsertLocation * columnCount: float * ?values: ResizeArray<ResizeArray<string>> -> unit
        /// <summary>Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
        /// <param name="rowCount">Required. Number of rows to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertRows: insertLocation: Word.InsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
        /// <param name="rowCount">Required. Number of rows to add.</param>
        /// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
        abstract insertRows: insertLocation: TableCellInsertRowsInsertLocation * rowCount: float * ?values: ResizeArray<ResizeArray<string>> -> Word.TableRowCollection
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: Word.CellPaddingLocation * cellPadding: float -> unit
        /// <summary>Sets cell padding in points.
        /// 
        /// [Api set: WordApi 1.3]</summary>
        /// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.</param>
        /// <param name="cellPadding">Required. The cell padding.</param>
        abstract setCellPadding: cellPaddingLocation: TableCellSetCellPaddingCellPaddingLocation * cellPadding: float -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.TableCellLoadOptions -> Word.TableCell
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableCell
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: TableCellLoadPropertyNamesAndPaths -> Word.TableCell
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableCell
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableCell
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.TableCell object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.TableCellData

    type [<StringEnum>] [<RequireQualifiedAccess>] TableCellGetBorderBorderLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right
        | [<CompiledName "InsideHorizontal">] InsideHorizontal
        | [<CompiledName "InsideVertical">] InsideVertical
        | [<CompiledName "Inside">] Inside
        | [<CompiledName "Outside">] Outside
        | [<CompiledName "All">] All

    type [<StringEnum>] [<RequireQualifiedAccess>] TableCellGetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<StringEnum>] [<RequireQualifiedAccess>] TableCellInsertColumnsInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableCellInsertRowsInsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] TableCellSetCellPaddingCellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<AllowNullLiteral>] TableCellLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Represents a table cell in a Word document.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCellStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableCell

    /// Contains the collection of the document's TableCell objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCellCollection =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets the loaded child items in this collection.
        abstract items: ResizeArray<Word.TableCell>
        /// Gets the first table cell in this collection. Throws an error if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirst: unit -> Word.TableCell
        /// Gets the first table cell in this collection. Returns a null object if this collection is empty.
        /// 
        /// [Api set: WordApi 1.3]
        abstract getFirstOrNullObject: unit -> Word.TableCell
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: obj -> Word.TableCellCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableCellCollection
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: OfficeExtension.LoadOption -> Word.TableCellCollection
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableCellCollection
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableCellCollection
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original `Word.TableCellCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        abstract toJSON: unit -> Word.Interfaces.TableCellCollectionData

    /// Contains the collection of the document's TableCell objects.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableCellCollectionStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableCellCollection

    /// Specifies the border style.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableBorder =
        inherit OfficeExtension.ClientObject
        /// The request context associated with the object. This connects the add-in's process to the Office host application's process.
        abstract context: RequestContext with get, set
        /// Gets or sets the table border color.
        /// 
        /// [Api set: WordApi 1.3]
        abstract color: string with get, set
        /// Gets or sets the type of the table border.
        /// 
        /// [Api set: WordApi 1.3]
        abstract ``type``: U2<Word.BorderType, string> with get, set
        /// Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
        /// 
        /// [Api set: WordApi 1.3]
        abstract width: float with get, set
        /// <summary>Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.</summary>
        /// <param name="properties">A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.</param>
        /// <param name="options">Provides an option to suppress errors if the properties object tries to set any read-only properties.</param>
        abstract set: properties: Interfaces.TableBorderUpdateData * ?options: OfficeExtension.UpdateOptions -> unit
        /// Sets multiple properties on the object at the same time, based on an existing loaded object.
        abstract set: properties: Word.TableBorder -> unit
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="options">Provides options for which properties of the object to load.</param>
        abstract load: ?options: Word.Interfaces.TableBorderLoadOptions -> Word.TableBorder
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNames">A comma-delimited string or an array of strings that specify the properties to load.</param>
        abstract load: ?propertyNames: U2<string, ResizeArray<string>> -> Word.TableBorder
        /// <summary>Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.</summary>
        /// <param name="propertyNamesAndPaths">`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.</param>
        abstract load: ?propertyNamesAndPaths: TableBorderLoadPropertyNamesAndPaths -> Word.TableBorder
        /// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for `context.trackedObjects.add(thisObject)`. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
        abstract track: unit -> Word.TableBorder
        /// Release the memory associated with this object, if it has previously been tracked. This call is shorthand for `context.trackedObjects.remove(thisObject)`. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
        abstract untrack: unit -> Word.TableBorder
        /// Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        /// Whereas the original Word.TableBorder object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableBorderData`) that contains shallow copies of any loaded child properties from the original object.
        abstract toJSON: unit -> Word.Interfaces.TableBorderData

    type [<AllowNullLiteral>] TableBorderLoadPropertyNamesAndPaths =
        abstract select: string option with get, set
        abstract expand: string option with get, set

    /// Specifies the border style.
    /// 
    /// [Api set: WordApi 1.3]
    type [<AllowNullLiteral>] TableBorderStatic =
        [<Emit "new $0($1...)">] abstract Create: unit -> TableBorder

    type [<StringEnum>] [<RequireQualifiedAccess>] EventType =
        | [<CompiledName "ContentControlDeleted">] ContentControlDeleted
        | [<CompiledName "ContentControlSelectionChanged">] ContentControlSelectionChanged
        | [<CompiledName "ContentControlDataChanged">] ContentControlDataChanged
        | [<CompiledName "ContentControlAdded">] ContentControlAdded
        | [<CompiledName "AnnotationAdded">] AnnotationAdded
        | [<CompiledName "AnnotationChanged">] AnnotationChanged
        | [<CompiledName "AnnotationDeleted">] AnnotationDeleted

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlType =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "RichTextInline">] RichTextInline
        | [<CompiledName "RichTextParagraphs">] RichTextParagraphs
        | [<CompiledName "RichTextTableCell">] RichTextTableCell
        | [<CompiledName "RichTextTableRow">] RichTextTableRow
        | [<CompiledName "RichTextTable">] RichTextTable
        | [<CompiledName "PlainTextInline">] PlainTextInline
        | [<CompiledName "PlainTextParagraph">] PlainTextParagraph
        | [<CompiledName "Picture">] Picture
        | [<CompiledName "BuildingBlockGallery">] BuildingBlockGallery
        | [<CompiledName "CheckBox">] CheckBox
        | [<CompiledName "ComboBox">] ComboBox
        | [<CompiledName "DropDownList">] DropDownList
        | [<CompiledName "DatePicker">] DatePicker
        | [<CompiledName "RepeatingSection">] RepeatingSection
        | [<CompiledName "RichText">] RichText
        | [<CompiledName "PlainText">] PlainText

    type [<StringEnum>] [<RequireQualifiedAccess>] ContentControlAppearance =
        | [<CompiledName "BoundingBox">] BoundingBox
        | [<CompiledName "Tags">] Tags2
        | [<CompiledName "Hidden">] Hidden

    type [<StringEnum>] [<RequireQualifiedAccess>] UnderlineType =
        | [<CompiledName "Mixed">] Mixed
        | [<CompiledName "None">] None
        | [<CompiledName "Hidden">] Hidden
        | [<CompiledName "DotLine">] DotLine
        | [<CompiledName "Single">] Single
        | [<CompiledName "Word">] Word
        | [<CompiledName "Double">] Double
        | [<CompiledName "Thick">] Thick
        | [<CompiledName "Dotted">] Dotted
        | [<CompiledName "DottedHeavy">] DottedHeavy
        | [<CompiledName "DashLine">] DashLine
        | [<CompiledName "DashLineHeavy">] DashLineHeavy
        | [<CompiledName "DashLineLong">] DashLineLong
        | [<CompiledName "DashLineLongHeavy">] DashLineLongHeavy
        | [<CompiledName "DotDashLine">] DotDashLine
        | [<CompiledName "DotDashLineHeavy">] DotDashLineHeavy
        | [<CompiledName "TwoDotDashLine">] TwoDotDashLine
        | [<CompiledName "TwoDotDashLineHeavy">] TwoDotDashLineHeavy
        | [<CompiledName "Wave">] Wave
        | [<CompiledName "WaveHeavy">] WaveHeavy
        | [<CompiledName "WaveDouble">] WaveDouble

    type [<StringEnum>] [<RequireQualifiedAccess>] BreakType =
        | [<CompiledName "Page">] Page
        | [<CompiledName "Next">] Next
        | [<CompiledName "SectionNext">] SectionNext
        | [<CompiledName "SectionContinuous">] SectionContinuous
        | [<CompiledName "SectionEven">] SectionEven
        | [<CompiledName "SectionOdd">] SectionOdd
        | [<CompiledName "Line">] Line

    type [<StringEnum>] [<RequireQualifiedAccess>] InsertLocation =
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Replace">] Replace

    type [<StringEnum>] [<RequireQualifiedAccess>] Alignment =
        | [<CompiledName "Mixed">] Mixed
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "Left">] Left
        | [<CompiledName "Centered">] Centered
        | [<CompiledName "Right">] Right
        | [<CompiledName "Justified">] Justified

    type [<StringEnum>] [<RequireQualifiedAccess>] HeaderFooterType =
        | [<CompiledName "Primary">] Primary
        | [<CompiledName "FirstPage">] FirstPage
        | [<CompiledName "EvenPages">] EvenPages

    type [<StringEnum>] [<RequireQualifiedAccess>] BodyType =
        | [<CompiledName "Unknown">] Unknown
        | [<CompiledName "MainDoc">] MainDoc
        | [<CompiledName "Section">] Section
        | [<CompiledName "Header">] Header
        | [<CompiledName "Footer">] Footer
        | [<CompiledName "TableCell">] TableCell

    type [<StringEnum>] [<RequireQualifiedAccess>] SelectionMode =
        | [<CompiledName "Select">] Select
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End

    type [<StringEnum>] [<RequireQualifiedAccess>] ImageFormat =
        | [<CompiledName "Unsupported">] Unsupported
        | [<CompiledName "Undefined">] Undefined
        | [<CompiledName "Bmp">] Bmp
        | [<CompiledName "Jpeg">] Jpeg
        | [<CompiledName "Gif">] Gif
        | [<CompiledName "Tiff">] Tiff
        | [<CompiledName "Png">] Png
        | [<CompiledName "Icon">] Icon
        | [<CompiledName "Exif">] Exif
        | [<CompiledName "Wmf">] Wmf
        | [<CompiledName "Emf">] Emf
        | [<CompiledName "Pict">] Pict
        | [<CompiledName "Pdf">] Pdf
        | [<CompiledName "Svg">] Svg

    type [<StringEnum>] [<RequireQualifiedAccess>] RangeLocation =
        | [<CompiledName "Whole">] Whole
        | [<CompiledName "Start">] Start
        | [<CompiledName "End">] End
        | [<CompiledName "Before">] Before
        | [<CompiledName "After">] After
        | [<CompiledName "Content">] Content

    type [<StringEnum>] [<RequireQualifiedAccess>] LocationRelation =
        | [<CompiledName "Unrelated">] Unrelated
        | [<CompiledName "Equal">] Equal
        | [<CompiledName "ContainsStart">] ContainsStart
        | [<CompiledName "ContainsEnd">] ContainsEnd
        | [<CompiledName "Contains">] Contains
        | [<CompiledName "InsideStart">] InsideStart
        | [<CompiledName "InsideEnd">] InsideEnd
        | [<CompiledName "Inside">] Inside
        | [<CompiledName "AdjacentBefore">] AdjacentBefore
        | [<CompiledName "OverlapsBefore">] OverlapsBefore
        | [<CompiledName "Before">] Before
        | [<CompiledName "AdjacentAfter">] AdjacentAfter
        | [<CompiledName "OverlapsAfter">] OverlapsAfter
        | [<CompiledName "After">] After

    type [<StringEnum>] [<RequireQualifiedAccess>] BorderLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right
        | [<CompiledName "InsideHorizontal">] InsideHorizontal
        | [<CompiledName "InsideVertical">] InsideVertical
        | [<CompiledName "Inside">] Inside
        | [<CompiledName "Outside">] Outside
        | [<CompiledName "All">] All

    type [<StringEnum>] [<RequireQualifiedAccess>] CellPaddingLocation =
        | [<CompiledName "Top">] Top
        | [<CompiledName "Left">] Left
        | [<CompiledName "Bottom">] Bottom
        | [<CompiledName "Right">] Right

    type [<StringEnum>] [<RequireQualifiedAccess>] BorderType =
        | [<CompiledName "Mixed">] Mixed
        | [<CompiledName "None">] None
        | [<CompiledName "Single">] Single
        | [<CompiledName "Double">] Double
        | [<CompiledName "Dotted">] Dotted
        | [<CompiledName "Dashed">] Dashed
        | [<CompiledName "DotDashed">] DotDashed
        | [<CompiledName "Dot2Dashed">] Dot2Dashed
        | [<CompiledName "Triple">] Triple
        | [<CompiledName "ThinThickSmall">] ThinThickSmall
        | [<CompiledName "ThickThinSmall">] ThickThinSmall
        | [<CompiledName "ThinThickThinSmall">] ThinThickThinSmall
        | [<CompiledName "ThinThickMed">] ThinThickMed
        | [<CompiledName "ThickThinMed">] ThickThinMed
        | [<CompiledName "ThinThickThinMed">] ThinThickThinMed
        | [<CompiledName "ThinThickLarge">] ThinThickLarge
        | [<CompiledName "ThickThinLarge">] ThickThinLarge
        | [<CompiledName "ThinThickThinLarge">] ThinThickThinLarge
        | [<CompiledName "Wave">] Wave
        | [<CompiledName "DoubleWave">] DoubleWave
        | [<CompiledName "DashedSmall">] DashedSmall
        | [<CompiledName "DashDotStroked">] DashDotStroked
        | [<CompiledName "ThreeDEmboss">] ThreeDEmboss
        | [<CompiledName "ThreeDEngrave">] ThreeDEngrave

    type [<StringEnum>] [<RequireQualifiedAccess>] VerticalAlignment =
        | [<CompiledName "Mixed">] Mixed
        | [<CompiledName "Top">] Top
        | [<CompiledName "Center">] Center
        | [<CompiledName "Bottom">] Bottom

    type [<StringEnum>] [<RequireQualifiedAccess>] ListLevelType =
        | [<CompiledName "Bullet">] Bullet
        | [<CompiledName "Number">] Number
        | [<CompiledName "Picture">] Picture

    type [<StringEnum>] [<RequireQualifiedAccess>] ListBullet =
        | [<CompiledName "Custom">] Custom
        | [<CompiledName "Solid">] Solid
        | [<CompiledName "Hollow">] Hollow
        | [<CompiledName "Square">] Square
        | [<CompiledName "Diamonds">] Diamonds
        | [<CompiledName "Arrow">] Arrow
        | [<CompiledName "Checkmark">] Checkmark

    type [<StringEnum>] [<RequireQualifiedAccess>] ListNumbering =
        | [<CompiledName "None">] None
        | [<CompiledName "Arabic">] Arabic
        | [<CompiledName "UpperRoman">] UpperRoman
        | [<CompiledName "LowerRoman">] LowerRoman
        | [<CompiledName "UpperLetter">] UpperLetter
        | [<CompiledName "LowerLetter">] LowerLetter

    type [<StringEnum>] [<RequireQualifiedAccess>] Style =
        | [<CompiledName "Other">] Other
        | [<CompiledName "Normal">] Normal
        | [<CompiledName "Heading1">] Heading1
        | [<CompiledName "Heading2">] Heading2
        | [<CompiledName "Heading3">] Heading3
        | [<CompiledName "Heading4">] Heading4
        | [<CompiledName "Heading5">] Heading5
        | [<CompiledName "Heading6">] Heading6
        | [<CompiledName "Heading7">] Heading7
        | [<CompiledName "Heading8">] Heading8
        | [<CompiledName "Heading9">] Heading9
        | [<CompiledName "Toc1">] Toc1
        | [<CompiledName "Toc2">] Toc2
        | [<CompiledName "Toc3">] Toc3
        | [<CompiledName "Toc4">] Toc4
        | [<CompiledName "Toc5">] Toc5
        | [<CompiledName "Toc6">] Toc6
        | [<CompiledName "Toc7">] Toc7
        | [<CompiledName "Toc8">] Toc8
        | [<CompiledName "Toc9">] Toc9
        | [<CompiledName "FootnoteText">] FootnoteText
        | [<CompiledName "Header">] Header
        | [<CompiledName "Footer">] Footer
        | [<CompiledName "Caption">] Caption
        | [<CompiledName "FootnoteReference">] FootnoteReference
        | [<CompiledName "EndnoteReference">] EndnoteReference
        | [<CompiledName "EndnoteText">] EndnoteText
        | [<CompiledName "Title">] Title
        | [<CompiledName "Subtitle">] Subtitle
        | [<CompiledName "Hyperlink">] Hyperlink
        | [<CompiledName "Strong">] Strong
        | [<CompiledName "Emphasis">] Emphasis
        | [<CompiledName "NoSpacing">] NoSpacing
        | [<CompiledName "ListParagraph">] ListParagraph
        | [<CompiledName "Quote">] Quote
        | [<CompiledName "IntenseQuote">] IntenseQuote
        | [<CompiledName "SubtleEmphasis">] SubtleEmphasis
        | [<CompiledName "IntenseEmphasis">] IntenseEmphasis
        | [<CompiledName "SubtleReference">] SubtleReference
        | [<CompiledName "IntenseReference">] IntenseReference
        | [<CompiledName "BookTitle">] BookTitle
        | [<CompiledName "Bibliography">] Bibliography
        | [<CompiledName "TocHeading">] TocHeading
        | [<CompiledName "TableGrid">] TableGrid
        | [<CompiledName "PlainTable1">] PlainTable1
        | [<CompiledName "PlainTable2">] PlainTable2
        | [<CompiledName "PlainTable3">] PlainTable3
        | [<CompiledName "PlainTable4">] PlainTable4
        | [<CompiledName "PlainTable5">] PlainTable5
        | [<CompiledName "TableGridLight">] TableGridLight
        | [<CompiledName "GridTable1Light">] GridTable1Light
        | [<CompiledName "GridTable1Light_Accent1">] GridTable1Light_Accent1
        | [<CompiledName "GridTable1Light_Accent2">] GridTable1Light_Accent2
        | [<CompiledName "GridTable1Light_Accent3">] GridTable1Light_Accent3
        | [<CompiledName "GridTable1Light_Accent4">] GridTable1Light_Accent4
        | [<CompiledName "GridTable1Light_Accent5">] GridTable1Light_Accent5
        | [<CompiledName "GridTable1Light_Accent6">] GridTable1Light_Accent6
        | [<CompiledName "GridTable2">] GridTable2
        | [<CompiledName "GridTable2_Accent1">] GridTable2_Accent1
        | [<CompiledName "GridTable2_Accent2">] GridTable2_Accent2
        | [<CompiledName "GridTable2_Accent3">] GridTable2_Accent3
        | [<CompiledName "GridTable2_Accent4">] GridTable2_Accent4
        | [<CompiledName "GridTable2_Accent5">] GridTable2_Accent5
        | [<CompiledName "GridTable2_Accent6">] GridTable2_Accent6
        | [<CompiledName "GridTable3">] GridTable3
        | [<CompiledName "GridTable3_Accent1">] GridTable3_Accent1
        | [<CompiledName "GridTable3_Accent2">] GridTable3_Accent2
        | [<CompiledName "GridTable3_Accent3">] GridTable3_Accent3
        | [<CompiledName "GridTable3_Accent4">] GridTable3_Accent4
        | [<CompiledName "GridTable3_Accent5">] GridTable3_Accent5
        | [<CompiledName "GridTable3_Accent6">] GridTable3_Accent6
        | [<CompiledName "GridTable4">] GridTable4
        | [<CompiledName "GridTable4_Accent1">] GridTable4_Accent1
        | [<CompiledName "GridTable4_Accent2">] GridTable4_Accent2
        | [<CompiledName "GridTable4_Accent3">] GridTable4_Accent3
        | [<CompiledName "GridTable4_Accent4">] GridTable4_Accent4
        | [<CompiledName "GridTable4_Accent5">] GridTable4_Accent5
        | [<CompiledName "GridTable4_Accent6">] GridTable4_Accent6
        | [<CompiledName "GridTable5Dark">] GridTable5Dark
        | [<CompiledName "GridTable5Dark_Accent1">] GridTable5Dark_Accent1
        | [<CompiledName "GridTable5Dark_Accent2">] GridTable5Dark_Accent2
        | [<CompiledName "GridTable5Dark_Accent3">] GridTable5Dark_Accent3
        | [<CompiledName "GridTable5Dark_Accent4">] GridTable5Dark_Accent4
        | [<CompiledName "GridTable5Dark_Accent5">] GridTable5Dark_Accent5
        | [<CompiledName "GridTable5Dark_Accent6">] GridTable5Dark_Accent6
        | [<CompiledName "GridTable6Colorful">] GridTable6Colorful
        | [<CompiledName "GridTable6Colorful_Accent1">] GridTable6Colorful_Accent1
        | [<CompiledName "GridTable6Colorful_Accent2">] GridTable6Colorful_Accent2
        | [<CompiledName "GridTable6Colorful_Accent3">] GridTable6Colorful_Accent3
        | [<CompiledName "GridTable6Colorful_Accent4">] GridTable6Colorful_Accent4
        | [<CompiledName "GridTable6Colorful_Accent5">] GridTable6Colorful_Accent5
        | [<CompiledName "GridTable6Colorful_Accent6">] GridTable6Colorful_Accent6
        | [<CompiledName "GridTable7Colorful">] GridTable7Colorful
        | [<CompiledName "GridTable7Colorful_Accent1">] GridTable7Colorful_Accent1
        | [<CompiledName "GridTable7Colorful_Accent2">] GridTable7Colorful_Accent2
        | [<CompiledName "GridTable7Colorful_Accent3">] GridTable7Colorful_Accent3
        | [<CompiledName "GridTable7Colorful_Accent4">] GridTable7Colorful_Accent4
        | [<CompiledName "GridTable7Colorful_Accent5">] GridTable7Colorful_Accent5
        | [<CompiledName "GridTable7Colorful_Accent6">] GridTable7Colorful_Accent6
        | [<CompiledName "ListTable1Light">] ListTable1Light
        | [<CompiledName "ListTable1Light_Accent1">] ListTable1Light_Accent1
        | [<CompiledName "ListTable1Light_Accent2">] ListTable1Light_Accent2
        | [<CompiledName "ListTable1Light_Accent3">] ListTable1Light_Accent3
        | [<CompiledName "ListTable1Light_Accent4">] ListTable1Light_Accent4
        | [<CompiledName "ListTable1Light_Accent5">] ListTable1Light_Accent5
        | [<CompiledName "ListTable1Light_Accent6">] ListTable1Light_Accent6
        | [<CompiledName "ListTable2">] ListTable2
        | [<CompiledName "ListTable2_Accent1">] ListTable2_Accent1
        | [<CompiledName "ListTable2_Accent2">] ListTable2_Accent2
        | [<CompiledName "ListTable2_Accent3">] ListTable2_Accent3
        | [<CompiledName "ListTable2_Accent4">] ListTable2_Accent4
        | [<CompiledName "ListTable2_Accent5">] ListTable2_Accent5
        | [<CompiledName "ListTable2_Accent6">] ListTable2_Accent6
        | [<CompiledName "ListTable3">] ListTable3
        | [<CompiledName "ListTable3_Accent1">] ListTable3_Accent1
        | [<CompiledName "ListTable3_Accent2">] ListTable3_Accent2
        | [<CompiledName "ListTable3_Accent3">] ListTable3_Accent3
        | [<CompiledName "ListTable3_Accent4">] ListTable3_Accent4
        | [<CompiledName "ListTable3_Accent5">] ListTable3_Accent5
        | [<CompiledName "ListTable3_Accent6">] ListTable3_Accent6
        | [<CompiledName "ListTable4">] ListTable4
        | [<CompiledName "ListTable4_Accent1">] ListTable4_Accent1
        | [<CompiledName "ListTable4_Accent2">] ListTable4_Accent2
        | [<CompiledName "ListTable4_Accent3">] ListTable4_Accent3
        | [<CompiledName "ListTable4_Accent4">] ListTable4_Accent4
        | [<CompiledName "ListTable4_Accent5">] ListTable4_Accent5
        | [<CompiledName "ListTable4_Accent6">] ListTable4_Accent6
        | [<CompiledName "ListTable5Dark">] ListTable5Dark
        | [<CompiledName "ListTable5Dark_Accent1">] ListTable5Dark_Accent1
        | [<CompiledName "ListTable5Dark_Accent2">] ListTable5Dark_Accent2
        | [<CompiledName "ListTable5Dark_Accent3">] ListTable5Dark_Accent3
        | [<CompiledName "ListTable5Dark_Accent4">] ListTable5Dark_Accent4
        | [<CompiledName "ListTable5Dark_Accent5">] ListTable5Dark_Accent5
        | [<CompiledName "ListTable5Dark_Accent6">] ListTable5Dark_Accent6
        | [<CompiledName "ListTable6Colorful">] ListTable6Colorful
        | [<CompiledName "ListTable6Colorful_Accent1">] ListTable6Colorful_Accent1
        | [<CompiledName "ListTable6Colorful_Accent2">] ListTable6Colorful_Accent2
        | [<CompiledName "ListTable6Colorful_Accent3">] ListTable6Colorful_Accent3
        | [<CompiledName "ListTable6Colorful_Accent4">] ListTable6Colorful_Accent4
        | [<CompiledName "ListTable6Colorful_Accent5">] ListTable6Colorful_Accent5
        | [<CompiledName "ListTable6Colorful_Accent6">] ListTable6Colorful_Accent6
        | [<CompiledName "ListTable7Colorful">] ListTable7Colorful
        | [<CompiledName "ListTable7Colorful_Accent1">] ListTable7Colorful_Accent1
        | [<CompiledName "ListTable7Colorful_Accent2">] ListTable7Colorful_Accent2
        | [<CompiledName "ListTable7Colorful_Accent3">] ListTable7Colorful_Accent3
        | [<CompiledName "ListTable7Colorful_Accent4">] ListTable7Colorful_Accent4
        | [<CompiledName "ListTable7Colorful_Accent5">] ListTable7Colorful_Accent5
        | [<CompiledName "ListTable7Colorful_Accent6">] ListTable7Colorful_Accent6

    type [<StringEnum>] [<RequireQualifiedAccess>] DocumentPropertyType =
        | [<CompiledName "String">] String
        | [<CompiledName "Number">] Number
        | [<CompiledName "Date">] Date
        | [<CompiledName "Boolean">] Boolean

    type [<StringEnum>] [<RequireQualifiedAccess>] TapObjectType =
        | [<CompiledName "Chart">] Chart
        | [<CompiledName "SmartArt">] SmartArt
        | [<CompiledName "Table">] Table
        | [<CompiledName "Image">] Image
        | [<CompiledName "Slide">] Slide
        | [<CompiledName "OLE">] Ole
        | [<CompiledName "Text">] Text

    type [<StringEnum>] [<RequireQualifiedAccess>] FileContentFormat =
        | [<CompiledName "Base64">] Base64
        | [<CompiledName "Html">] Html
        | [<CompiledName "Ooxml">] Ooxml

    type [<StringEnum>] [<RequireQualifiedAccess>] ErrorCodes =
        | [<CompiledName "AccessDenied">] AccessDenied
        | [<CompiledName "GeneralException">] GeneralException
        | [<CompiledName "InvalidArgument">] InvalidArgument
        | [<CompiledName "ItemNotFound">] ItemNotFound
        | [<CompiledName "NotImplemented">] NotImplemented
        | [<CompiledName "SearchDialogIsOpen">] SearchDialogIsOpen
        | [<CompiledName "SearchStringInvalidOrTooLong">] SearchStringInvalidOrTooLong

    module Interfaces =

        /// Provides ways to load properties of only a subset of members of a collection.
        type [<AllowNullLiteral>] CollectionLoadOptions =
            /// Specify the number of items in the queried collection to be included in the result.
            abstract ``$top``: float option with get, set
            /// Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            abstract ``$skip``: float option with get, set

        /// An interface for updating data on the Body object, for use in `body.set({ ... })`.
        type [<AllowNullLiteral>] BodyUpdateData =
            /// Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set

        /// An interface for updating data on the ContentControl object, for use in `contentControl.set({ ... })`.
        type [<AllowNullLiteral>] ContentControlUpdateData =
            /// Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract appearance: U2<Word.ContentControlAppearance, string> option with get, set
            /// Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotDelete: bool option with get, set
            /// Gets or sets a value that indicates whether the user can edit the contents of the content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotEdit: bool option with get, set
            /// Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: string option with get, set
            /// Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
            /// 
            /// **Note**: The set operation for this property is not supported in Word on the web.
            /// 
            /// [Api set: WordApi 1.1]
            abstract placeholderText: string option with get, set
            /// Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
            /// 
            /// [Api set: WordApi 1.1]
            abstract removeWhenEdited: bool option with get, set
            /// Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets or sets a tag to identify a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract tag: string option with get, set
            /// Gets or sets the title for a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract title: string option with get, set

        /// An interface for updating data on the ContentControlCollection object, for use in `contentControlCollection.set({ ... })`.
        type [<AllowNullLiteral>] ContentControlCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.ContentControlData> option with get, set

        /// An interface for updating data on the CustomProperty object, for use in `customProperty.set({ ... })`.
        type [<AllowNullLiteral>] CustomPropertyUpdateData =
            /// Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: obj option with get, set

        /// An interface for updating data on the CustomPropertyCollection object, for use in `customPropertyCollection.set({ ... })`.
        type [<AllowNullLiteral>] CustomPropertyCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.CustomPropertyData> option with get, set

        /// An interface for updating data on the Document object, for use in `document.set({ ... })`.
        type [<AllowNullLiteral>] DocumentUpdateData =
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyUpdateData option with get, set
            /// Gets the properties of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesUpdateData option with get, set

        /// An interface for updating data on the DocumentCreated object, for use in `documentCreated.set({ ... })`.
        type [<AllowNullLiteral>] DocumentCreatedUpdateData =
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract body: Word.Interfaces.BodyUpdateData option with get, set
            /// Gets the properties of the document.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesUpdateData option with get, set

        /// An interface for updating data on the DocumentProperties object, for use in `documentProperties.set({ ... })`.
        type [<AllowNullLiteral>] DocumentPropertiesUpdateData =
            /// Gets or sets the author of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract author: string option with get, set
            /// Gets or sets the category of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract category: string option with get, set
            /// Gets or sets the comments of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract comments: string option with get, set
            /// Gets or sets the company of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract company: string option with get, set
            /// Gets or sets the format of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract format: string option with get, set
            /// Gets or sets the keywords of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract keywords: string option with get, set
            /// Gets or sets the manager of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract manager: string option with get, set
            /// Gets or sets the subject of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subject: string option with get, set
            /// Gets or sets the title of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract title: string option with get, set

        /// An interface for updating data on the Font object, for use in `font.set({ ... })`.
        type [<AllowNullLiteral>] FontUpdateData =
            /// Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract bold: bool option with get, set
            /// Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: string option with get, set
            /// Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract doubleStrikeThrough: bool option with get, set
            /// Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
            ///           *Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract highlightColor: string option with get, set
            /// Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract italic: bool option with get, set
            /// Gets or sets a value that represents the name of the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract name: string option with get, set
            /// Gets or sets a value that represents the font size in points.
            /// 
            /// [Api set: WordApi 1.1]
            abstract size: float option with get, set
            /// Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract strikeThrough: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract subscript: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract superscript: bool option with get, set
            /// Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
            /// 
            /// [Api set: WordApi 1.1]
            abstract underline: U2<Word.UnderlineType, string> option with get, set

        /// An interface for updating data on the InlinePicture object, for use in `inlinePicture.set({ ... })`.
        type [<AllowNullLiteral>] InlinePictureUpdateData =
            /// Gets or sets a string that represents the alternative text associated with the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextDescription: string option with get, set
            /// Gets or sets a string that contains the title for the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextTitle: string option with get, set
            /// Gets or sets a number that describes the height of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract height: float option with get, set
            /// Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.1]
            abstract hyperlink: string option with get, set
            /// Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lockAspectRatio: bool option with get, set
            /// Gets or sets a number that describes the width of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract width: float option with get, set

        /// An interface for updating data on the InlinePictureCollection object, for use in `inlinePictureCollection.set({ ... })`.
        type [<AllowNullLiteral>] InlinePictureCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set

        /// An interface for updating data on the ListCollection object, for use in `listCollection.set({ ... })`.
        type [<AllowNullLiteral>] ListCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.ListData> option with get, set

        /// An interface for updating data on the ListItem object, for use in `listItem.set({ ... })`.
        type [<AllowNullLiteral>] ListItemUpdateData =
            /// Gets or sets the level of the item in the list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract level: float option with get, set

        /// An interface for updating data on the Paragraph object, for use in `paragraph.set({ ... })`.
        type [<AllowNullLiteral>] ParagraphUpdateData =
            /// Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItem: Word.Interfaces.ListItemUpdateData option with get, set
            /// Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItemOrNullObject: Word.Interfaces.ListItemUpdateData option with get, set
            /// Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract alignment: U2<Word.Alignment, string> option with get, set
            /// Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
            /// 
            /// [Api set: WordApi 1.1]
            abstract firstLineIndent: float option with get, set
            /// Gets or sets the left indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract leftIndent: float option with get, set
            /// Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineSpacing: float option with get, set
            /// Gets or sets the amount of spacing, in grid lines, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitAfter: float option with get, set
            /// Gets or sets the amount of spacing, in grid lines, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitBefore: float option with get, set
            /// Gets or sets the outline level for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract outlineLevel: float option with get, set
            /// Gets or sets the right indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract rightIndent: float option with get, set
            /// Gets or sets the spacing, in points, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceAfter: float option with get, set
            /// Gets or sets the spacing, in points, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceBefore: float option with get, set
            /// Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set

        /// An interface for updating data on the ParagraphCollection object, for use in `paragraphCollection.set({ ... })`.
        type [<AllowNullLiteral>] ParagraphCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.ParagraphData> option with get, set

        /// An interface for updating data on the Range object, for use in `range.set({ ... })`.
        type [<AllowNullLiteral>] RangeUpdateData =
            /// Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.3]
            abstract hyperlink: string option with get, set
            /// Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set

        /// An interface for updating data on the RangeCollection object, for use in `rangeCollection.set({ ... })`.
        type [<AllowNullLiteral>] RangeCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.RangeData> option with get, set

        /// An interface for updating data on the SearchOptions object, for use in `searchOptions.set({ ... })`.
        type [<AllowNullLiteral>] SearchOptionsUpdateData =
            /// Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignorePunct: bool option with get, set
            /// Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignoreSpace: bool option with get, set
            /// Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchCase: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchPrefix: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchSuffix: bool option with get, set
            /// Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWholeWord: bool option with get, set
            /// Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWildcards: bool option with get, set

        /// An interface for updating data on the Section object, for use in `section.set({ ... })`.
        type [<AllowNullLiteral>] SectionUpdateData =
            /// Gets the body object of the section. This does not include the header/footer and other section metadata.
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyUpdateData option with get, set

        /// An interface for updating data on the SectionCollection object, for use in `sectionCollection.set({ ... })`.
        type [<AllowNullLiteral>] SectionCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.SectionData> option with get, set

        /// An interface for updating data on the Table object, for use in `table.set({ ... })`.
        type [<AllowNullLiteral>] TableUpdateData =
            /// Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract alignment: U2<Word.Alignment, string> option with get, set
            /// Gets and sets the number of header rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract headerRowCount: float option with get, set
            /// Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract style: string option with get, set
            /// Gets and sets whether the table has banded columns.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedColumns: bool option with get, set
            /// Gets and sets whether the table has banded rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedRows: bool option with get, set
            /// Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets and sets whether the table has a first column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleFirstColumn: bool option with get, set
            /// Gets and sets whether the table has a last column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleLastColumn: bool option with get, set
            /// Gets and sets whether the table has a total (last) row with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleTotalRow: bool option with get, set
            /// Gets and sets the text values in the table, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: ResizeArray<ResizeArray<string>> option with get, set
            /// Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set
            /// Gets and sets the width of the table in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: float option with get, set

        /// An interface for updating data on the TableCollection object, for use in `tableCollection.set({ ... })`.
        type [<AllowNullLiteral>] TableCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.TableData> option with get, set

        /// An interface for updating data on the TableRow object, for use in `tableRow.set({ ... })`.
        type [<AllowNullLiteral>] TableRowUpdateData =
            /// Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontUpdateData option with get, set
            /// Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Gets and sets the preferred height of the row in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract preferredHeight: float option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets and sets the text values in the row, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: ResizeArray<ResizeArray<string>> option with get, set
            /// Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set

        /// An interface for updating data on the TableRowCollection object, for use in `tableRowCollection.set({ ... })`.
        type [<AllowNullLiteral>] TableRowCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.TableRowData> option with get, set

        /// An interface for updating data on the TableCell object, for use in `tableCell.set({ ... })`.
        type [<AllowNullLiteral>] TableCellUpdateData =
            /// Gets the body object of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract body: Word.Interfaces.BodyUpdateData option with get, set
            /// Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
            /// 
            /// [Api set: WordApi 1.3]
            abstract columnWidth: float option with get, set
            /// Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets and sets the text of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: string option with get, set
            /// Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set

        /// An interface for updating data on the TableCellCollection object, for use in `tableCellCollection.set({ ... })`.
        type [<AllowNullLiteral>] TableCellCollectionUpdateData =
            abstract items: ResizeArray<Word.Interfaces.TableCellData> option with get, set

        /// An interface for updating data on the TableBorder object, for use in `tableBorder.set({ ... })`.
        type [<AllowNullLiteral>] TableBorderUpdateData =
            /// Gets or sets the table border color.
            /// 
            /// [Api set: WordApi 1.3]
            abstract color: string option with get, set
            /// Gets or sets the type of the table border.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: U2<Word.BorderType, string> option with get, set
            /// Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: float option with get, set

        /// An interface describing the data returned by calling `body.toJSON()`.
        type [<AllowNullLiteral>] BodyData =
            /// Gets the collection of rich text content control objects in the body. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract contentControls: ResizeArray<Word.Interfaces.ContentControlData> option with get, set
            /// Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract inlinePictures: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set
            /// Gets the collection of list objects in the body. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lists: ResizeArray<Word.Interfaces.ListData> option with get, set
            /// Gets the collection of paragraph objects in the body. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract paragraphs: ResizeArray<Word.Interfaces.ParagraphData> option with get, set
            /// Gets the collection of table objects in the body. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tables: ResizeArray<Word.Interfaces.TableData> option with get, set
            /// Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets the text of the body. Use the insertText method to insert text. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: string option with get, set
            /// Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: U2<Word.BodyType, string> option with get, set

        /// An interface describing the data returned by calling `contentControl.toJSON()`.
        type [<AllowNullLiteral>] ContentControlData =
            /// Gets the collection of content control objects in the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract contentControls: ResizeArray<Word.Interfaces.ContentControlData> option with get, set
            /// Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract inlinePictures: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set
            /// Gets the collection of list objects in the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lists: ResizeArray<Word.Interfaces.ListData> option with get, set
            /// Get the collection of paragraph objects in the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract paragraphs: ResizeArray<Word.Interfaces.ParagraphData> option with get, set
            /// Gets the collection of table objects in the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tables: ResizeArray<Word.Interfaces.TableData> option with get, set
            /// Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract appearance: U2<Word.ContentControlAppearance, string> option with get, set
            /// Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotDelete: bool option with get, set
            /// Gets or sets a value that indicates whether the user can edit the contents of the content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotEdit: bool option with get, set
            /// Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: string option with get, set
            /// Gets an integer that represents the content control identifier. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract id: float option with get, set
            /// Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
            /// 
            /// **Note**: The set operation for this property is not supported in Word on the web.
            /// 
            /// [Api set: WordApi 1.1]
            abstract placeholderText: string option with get, set
            /// Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
            /// 
            /// [Api set: WordApi 1.1]
            abstract removeWhenEdited: bool option with get, set
            /// Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subtype: U2<Word.ContentControlType, string> option with get, set
            /// Gets or sets a tag to identify a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract tag: string option with get, set
            /// Gets the text of the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: string option with get, set
            /// Gets or sets the title for a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract title: string option with get, set
            /// Gets the content control type. Only rich text content controls are supported currently. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ``type``: U2<Word.ContentControlType, string> option with get, set

        /// An interface describing the data returned by calling `contentControlCollection.toJSON()`.
        type [<AllowNullLiteral>] ContentControlCollectionData =
            abstract items: ResizeArray<Word.Interfaces.ContentControlData> option with get, set

        /// An interface describing the data returned by calling `customProperty.toJSON()`.
        type [<AllowNullLiteral>] CustomPropertyData =
            /// Gets the key of the custom property. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract key: string option with get, set
            /// Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: U2<Word.DocumentPropertyType, string> option with get, set
            /// Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: obj option with get, set

        /// An interface describing the data returned by calling `customPropertyCollection.toJSON()`.
        type [<AllowNullLiteral>] CustomPropertyCollectionData =
            abstract items: ResizeArray<Word.Interfaces.CustomPropertyData> option with get, set

        /// An interface describing the data returned by calling `document.toJSON()`.
        type [<AllowNullLiteral>] DocumentData =
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyData option with get, set
            /// Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract contentControls: ResizeArray<Word.Interfaces.ContentControlData> option with get, set
            /// Gets the properties of the document. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesData option with get, set
            /// Gets the collection of section objects in the document. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract sections: ResizeArray<Word.Interfaces.SectionData> option with get, set
            /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract saved: bool option with get, set

        /// An interface describing the data returned by calling `documentCreated.toJSON()`.
        type [<AllowNullLiteral>] DocumentCreatedData =
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract body: Word.Interfaces.BodyData option with get, set
            /// Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract contentControls: ResizeArray<Word.Interfaces.ContentControlData> option with get, set
            /// Gets the properties of the document. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesData option with get, set
            /// Gets the collection of section objects in the document. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract sections: ResizeArray<Word.Interfaces.SectionData> option with get, set
            /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract saved: bool option with get, set

        /// An interface describing the data returned by calling `documentProperties.toJSON()`.
        type [<AllowNullLiteral>] DocumentPropertiesData =
            /// Gets the collection of custom properties of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract customProperties: ResizeArray<Word.Interfaces.CustomPropertyData> option with get, set
            /// Gets the application name of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract applicationName: string option with get, set
            /// Gets or sets the author of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract author: string option with get, set
            /// Gets or sets the category of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract category: string option with get, set
            /// Gets or sets the comments of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract comments: string option with get, set
            /// Gets or sets the company of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract company: string option with get, set
            /// Gets the creation date of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract creationDate: DateTime option with get, set
            /// Gets or sets the format of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract format: string option with get, set
            /// Gets or sets the keywords of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract keywords: string option with get, set
            /// Gets the last author of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastAuthor: string option with get, set
            /// Gets the last print date of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastPrintDate: DateTime option with get, set
            /// Gets the last save time of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastSaveTime: DateTime option with get, set
            /// Gets or sets the manager of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract manager: string option with get, set
            /// Gets the revision number of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract revisionNumber: string option with get, set
            /// Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
            /// 
            /// [Api set: WordApi 1.3]
            abstract security: float option with get, set
            /// Gets or sets the subject of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subject: string option with get, set
            /// Gets the template of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract template: string option with get, set
            /// Gets or sets the title of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract title: string option with get, set

        /// An interface describing the data returned by calling `font.toJSON()`.
        type [<AllowNullLiteral>] FontData =
            /// Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract bold: bool option with get, set
            /// Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: string option with get, set
            /// Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract doubleStrikeThrough: bool option with get, set
            /// Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
            ///           *Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract highlightColor: string option with get, set
            /// Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract italic: bool option with get, set
            /// Gets or sets a value that represents the name of the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract name: string option with get, set
            /// Gets or sets a value that represents the font size in points.
            /// 
            /// [Api set: WordApi 1.1]
            abstract size: float option with get, set
            /// Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract strikeThrough: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract subscript: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract superscript: bool option with get, set
            /// Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
            /// 
            /// [Api set: WordApi 1.1]
            abstract underline: U2<Word.UnderlineType, string> option with get, set

        /// An interface describing the data returned by calling `inlinePicture.toJSON()`.
        type [<AllowNullLiteral>] InlinePictureData =
            /// Gets or sets a string that represents the alternative text associated with the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextDescription: string option with get, set
            /// Gets or sets a string that contains the title for the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextTitle: string option with get, set
            /// Gets or sets a number that describes the height of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract height: float option with get, set
            /// Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.1]
            abstract hyperlink: string option with get, set
            /// Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lockAspectRatio: bool option with get, set
            /// Gets or sets a number that describes the width of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract width: float option with get, set

        /// An interface describing the data returned by calling `inlinePictureCollection.toJSON()`.
        type [<AllowNullLiteral>] InlinePictureCollectionData =
            abstract items: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set

        /// An interface describing the data returned by calling `list.toJSON()`.
        type [<AllowNullLiteral>] ListData =
            /// Gets paragraphs in the list. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract paragraphs: ResizeArray<Word.Interfaces.ParagraphData> option with get, set
            /// Gets the list's id.
            /// 
            /// [Api set: WordApi 1.3]
            abstract id: float option with get, set
            /// Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelExistences: ResizeArray<bool> option with get, set
            /// Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelTypes: ResizeArray<Word.ListLevelType> option with get, set

        /// An interface describing the data returned by calling `listCollection.toJSON()`.
        type [<AllowNullLiteral>] ListCollectionData =
            abstract items: ResizeArray<Word.Interfaces.ListData> option with get, set

        /// An interface describing the data returned by calling `listItem.toJSON()`.
        type [<AllowNullLiteral>] ListItemData =
            /// Gets or sets the level of the item in the list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract level: float option with get, set
            /// Gets the list item bullet, number, or picture as a string. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listString: string option with get, set
            /// Gets the list item order number in relation to its siblings. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract siblingIndex: float option with get, set

        /// An interface describing the data returned by calling `paragraph.toJSON()`.
        type [<AllowNullLiteral>] ParagraphData =
            /// Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract inlinePictures: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set
            /// Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItem: Word.Interfaces.ListItemData option with get, set
            /// Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItemOrNullObject: Word.Interfaces.ListItemData option with get, set
            /// Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract alignment: U2<Word.Alignment, string> option with get, set
            /// Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
            /// 
            /// [Api set: WordApi 1.1]
            abstract firstLineIndent: float option with get, set
            /// Indicates the paragraph is the last one inside its parent body. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isLastParagraph: bool option with get, set
            /// Checks whether the paragraph is a list item. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isListItem: bool option with get, set
            /// Gets or sets the left indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract leftIndent: float option with get, set
            /// Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineSpacing: float option with get, set
            /// Gets or sets the amount of spacing, in grid lines, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitAfter: float option with get, set
            /// Gets or sets the amount of spacing, in grid lines, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitBefore: float option with get, set
            /// Gets or sets the outline level for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract outlineLevel: float option with get, set
            /// Gets or sets the right indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract rightIndent: float option with get, set
            /// Gets or sets the spacing, in points, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceAfter: float option with get, set
            /// Gets or sets the spacing, in points, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceBefore: float option with get, set
            /// Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tableNestingLevel: float option with get, set
            /// Gets the text of the paragraph. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: string option with get, set

        /// An interface describing the data returned by calling `paragraphCollection.toJSON()`.
        type [<AllowNullLiteral>] ParagraphCollectionData =
            abstract items: ResizeArray<Word.Interfaces.ParagraphData> option with get, set

        /// An interface describing the data returned by calling `range.toJSON()`.
        type [<AllowNullLiteral>] RangeData =
            /// Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets the collection of inline picture objects in the range. Read-only.
            /// 
            /// [Api set: WordApi 1.2]
            abstract inlinePictures: ResizeArray<Word.Interfaces.InlinePictureData> option with get, set
            /// Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.3]
            abstract hyperlink: string option with get, set
            /// Checks whether the range length is zero. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isEmpty: bool option with get, set
            /// Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: string option with get, set
            /// Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets the text of the range. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: string option with get, set

        /// An interface describing the data returned by calling `rangeCollection.toJSON()`.
        type [<AllowNullLiteral>] RangeCollectionData =
            abstract items: ResizeArray<Word.Interfaces.RangeData> option with get, set

        /// An interface describing the data returned by calling `searchOptions.toJSON()`.
        type [<AllowNullLiteral>] SearchOptionsData =
            /// Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignorePunct: bool option with get, set
            /// Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignoreSpace: bool option with get, set
            /// Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchCase: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchPrefix: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchSuffix: bool option with get, set
            /// Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWholeWord: bool option with get, set
            /// Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWildcards: bool option with get, set

        /// An interface describing the data returned by calling `section.toJSON()`.
        type [<AllowNullLiteral>] SectionData =
            /// Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyData option with get, set

        /// An interface describing the data returned by calling `sectionCollection.toJSON()`.
        type [<AllowNullLiteral>] SectionCollectionData =
            abstract items: ResizeArray<Word.Interfaces.SectionData> option with get, set

        /// An interface describing the data returned by calling `table.toJSON()`.
        type [<AllowNullLiteral>] TableData =
            /// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets all of the table rows. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rows: ResizeArray<Word.Interfaces.TableRowData> option with get, set
            /// Gets the child tables nested one level deeper. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tables: ResizeArray<Word.Interfaces.TableData> option with get, set
            /// Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract alignment: U2<Word.Alignment, string> option with get, set
            /// Gets and sets the number of header rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract headerRowCount: float option with get, set
            /// Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Indicates whether all of the table rows are uniform. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isUniform: bool option with get, set
            /// Gets the nesting level of the table. Top-level tables have level 1. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract nestingLevel: float option with get, set
            /// Gets the number of rows in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowCount: float option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract style: string option with get, set
            /// Gets and sets whether the table has banded columns.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedColumns: bool option with get, set
            /// Gets and sets whether the table has banded rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedRows: bool option with get, set
            /// Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: U2<Word.Style, string> option with get, set
            /// Gets and sets whether the table has a first column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleFirstColumn: bool option with get, set
            /// Gets and sets whether the table has a last column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleLastColumn: bool option with get, set
            /// Gets and sets whether the table has a total (last) row with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleTotalRow: bool option with get, set
            /// Gets and sets the text values in the table, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: ResizeArray<ResizeArray<string>> option with get, set
            /// Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set
            /// Gets and sets the width of the table in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: float option with get, set

        /// An interface describing the data returned by calling `tableCollection.toJSON()`.
        type [<AllowNullLiteral>] TableCollectionData =
            abstract items: ResizeArray<Word.Interfaces.TableData> option with get, set

        /// An interface describing the data returned by calling `tableRow.toJSON()`.
        type [<AllowNullLiteral>] TableRowData =
            /// Gets cells. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cells: ResizeArray<Word.Interfaces.TableCellData> option with get, set
            /// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontData option with get, set
            /// Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellCount: float option with get, set
            /// Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isHeader: bool option with get, set
            /// Gets and sets the preferred height of the row in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract preferredHeight: float option with get, set
            /// Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: float option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets and sets the text values in the row, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: ResizeArray<ResizeArray<string>> option with get, set
            /// Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set

        /// An interface describing the data returned by calling `tableRowCollection.toJSON()`.
        type [<AllowNullLiteral>] TableRowCollectionData =
            abstract items: ResizeArray<Word.Interfaces.TableRowData> option with get, set

        /// An interface describing the data returned by calling `tableCell.toJSON()`.
        type [<AllowNullLiteral>] TableCellData =
            /// Gets the body object of the cell. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract body: Word.Interfaces.BodyData option with get, set
            /// Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellIndex: float option with get, set
            /// Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
            /// 
            /// [Api set: WordApi 1.3]
            abstract columnWidth: float option with get, set
            /// Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: U2<Word.Alignment, string> option with get, set
            /// Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: float option with get, set
            /// Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: string option with get, set
            /// Gets and sets the text of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: string option with get, set
            /// Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: U2<Word.VerticalAlignment, string> option with get, set
            /// Gets the width of the cell in points. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: float option with get, set

        /// An interface describing the data returned by calling `tableCellCollection.toJSON()`.
        type [<AllowNullLiteral>] TableCellCollectionData =
            abstract items: ResizeArray<Word.Interfaces.TableCellData> option with get, set

        /// An interface describing the data returned by calling `tableBorder.toJSON()`.
        type [<AllowNullLiteral>] TableBorderData =
            /// Gets or sets the table border color.
            /// 
            /// [Api set: WordApi 1.3]
            abstract color: string option with get, set
            /// Gets or sets the type of the table border.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: U2<Word.BorderType, string> option with get, set
            /// Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: float option with get, set

        /// Represents the body of a document or a section.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] BodyLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an error if there isn't a parent body.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBodyOrNullObject: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the content control that contains the body. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the body. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the parent section of the body. Throws an error if there isn't a parent section.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentSection: Word.Interfaces.SectionLoadOptions option with get, set
            /// Gets the parent section of the body. Returns a null object if there isn't a parent section.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentSectionOrNullObject: Word.Interfaces.SectionLoadOptions option with get, set
            /// Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// Gets the text of the body. Use the insertText method to insert text. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set
            /// Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: bool option with get, set

        /// Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] ContentControlLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets the parent body of the content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the content control. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the table that contains the content control. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table that contains the content control. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract appearance: bool option with get, set
            /// Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotDelete: bool option with get, set
            /// Gets or sets a value that indicates whether the user can edit the contents of the content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotEdit: bool option with get, set
            /// Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: bool option with get, set
            /// Gets an integer that represents the content control identifier. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract id: bool option with get, set
            /// Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
            /// 
            /// **Note**: The set operation for this property is not supported in Word on the web.
            /// 
            /// [Api set: WordApi 1.1]
            abstract placeholderText: bool option with get, set
            /// Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
            /// 
            /// [Api set: WordApi 1.1]
            abstract removeWhenEdited: bool option with get, set
            /// Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subtype: bool option with get, set
            /// Gets or sets a tag to identify a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract tag: bool option with get, set
            /// Gets the text of the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set
            /// Gets or sets the title for a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract title: bool option with get, set
            /// Gets the content control type. Only rich text content controls are supported currently. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ``type``: bool option with get, set

        /// Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] ContentControlCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent body of the content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the content control. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the content control. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the content control. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract appearance: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotDelete: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can edit the contents of the content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract cannotEdit: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: bool option with get, set
            /// For EACH ITEM in the collection: Gets an integer that represents the content control identifier. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
            /// 
            /// **Note**: The set operation for this property is not supported in Word on the web.
            /// 
            /// [Api set: WordApi 1.1]
            abstract placeholderText: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
            /// 
            /// [Api set: WordApi 1.1]
            abstract removeWhenEdited: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// For EACH ITEM in the collection: Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subtype: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a tag to identify a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract tag: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text of the content control. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the title for a content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract title: bool option with get, set
            /// For EACH ITEM in the collection: Gets the content control type. Only rich text content controls are supported currently. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ``type``: bool option with get, set

        /// Represents a custom property.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] CustomPropertyLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the key of the custom property. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract key: bool option with get, set
            /// Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: bool option with get, set
            /// Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: bool option with get, set

        /// Contains the collection of {@link Word.CustomProperty} objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] CustomPropertyCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the key of the custom property. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract key: bool option with get, set
            /// For EACH ITEM in the collection: Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: bool option with get, set

        /// The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] DocumentLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the properties of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesLoadOptions option with get, set
            /// Gets or sets a value that indicates that, when opening a new document, whether it is allowed to close this document even if this document is untitled. True to close, false otherwise.
            /// 
            /// [Api set: WordApi]
            abstract allowCloseOnUntitled: bool option with get, set
            /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract saved: bool option with get, set

        /// The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] DocumentCreatedLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the properties of the document.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract properties: Word.Interfaces.DocumentPropertiesLoadOptions option with get, set
            /// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
            /// 
            /// [Api set: WordApiHiddenDocument 1.3]
            abstract saved: bool option with get, set

        /// Represents document properties.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] DocumentPropertiesLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the application name of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract applicationName: bool option with get, set
            /// Gets or sets the author of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract author: bool option with get, set
            /// Gets or sets the category of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract category: bool option with get, set
            /// Gets or sets the comments of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract comments: bool option with get, set
            /// Gets or sets the company of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract company: bool option with get, set
            /// Gets the creation date of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract creationDate: bool option with get, set
            /// Gets or sets the format of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract format: bool option with get, set
            /// Gets or sets the keywords of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract keywords: bool option with get, set
            /// Gets the last author of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastAuthor: bool option with get, set
            /// Gets the last print date of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastPrintDate: bool option with get, set
            /// Gets the last save time of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract lastSaveTime: bool option with get, set
            /// Gets or sets the manager of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract manager: bool option with get, set
            /// Gets the revision number of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract revisionNumber: bool option with get, set
            /// Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
            /// 
            /// [Api set: WordApi 1.3]
            abstract security: bool option with get, set
            /// Gets or sets the subject of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract subject: bool option with get, set
            /// Gets the template of the document. Read only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract template: bool option with get, set
            /// Gets or sets the title of the document.
            /// 
            /// [Api set: WordApi 1.3]
            abstract title: bool option with get, set

        /// Represents a font.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] FontLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract bold: bool option with get, set
            /// Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
            /// 
            /// [Api set: WordApi 1.1]
            abstract color: bool option with get, set
            /// Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract doubleStrikeThrough: bool option with get, set
            /// Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
            ///           *Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract highlightColor: bool option with get, set
            /// Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract italic: bool option with get, set
            /// Gets or sets a value that represents the name of the font.
            /// 
            /// [Api set: WordApi 1.1]
            abstract name: bool option with get, set
            /// Gets or sets a value that represents the font size in points.
            /// 
            /// [Api set: WordApi 1.1]
            abstract size: bool option with get, set
            /// Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract strikeThrough: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract subscript: bool option with get, set
            /// Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
            /// 
            /// [Api set: WordApi 1.1]
            abstract superscript: bool option with get, set
            /// Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
            /// 
            /// [Api set: WordApi 1.1]
            abstract underline: bool option with get, set

        /// Represents an inline picture.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] InlinePictureLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the parent paragraph that contains the inline image.
            /// 
            /// [Api set: WordApi 1.2]
            abstract paragraph: Word.Interfaces.ParagraphLoadOptions option with get, set
            /// Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the table that contains the inline image. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table that contains the inline image. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets or sets a string that represents the alternative text associated with the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextDescription: bool option with get, set
            /// Gets or sets a string that contains the title for the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextTitle: bool option with get, set
            /// Gets or sets a number that describes the height of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract height: bool option with get, set
            /// Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.1]
            abstract hyperlink: bool option with get, set
            /// Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lockAspectRatio: bool option with get, set
            /// Gets or sets a number that describes the width of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract width: bool option with get, set

        /// Contains a collection of {@link Word.InlinePicture} objects.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] InlinePictureCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the parent paragraph that contains the inline image.
            /// 
            /// [Api set: WordApi 1.2]
            abstract paragraph: Word.Interfaces.ParagraphLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the inline image. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the inline image. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets or sets a string that represents the alternative text associated with the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextDescription: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a string that contains the title for the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract altTextTitle: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a number that describes the height of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract height: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.1]
            abstract hyperlink: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lockAspectRatio: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets a number that describes the width of the inline image.
            /// 
            /// [Api set: WordApi 1.1]
            abstract width: bool option with get, set

        /// Contains a collection of {@link Word.Paragraph} objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] ListLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the list's id.
            /// 
            /// [Api set: WordApi 1.3]
            abstract id: bool option with get, set
            /// Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelExistences: bool option with get, set
            /// Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelTypes: bool option with get, set

        /// Contains a collection of {@link Word.List} objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] ListCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the list's id.
            /// 
            /// [Api set: WordApi 1.3]
            abstract id: bool option with get, set
            /// For EACH ITEM in the collection: Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelExistences: bool option with get, set
            /// For EACH ITEM in the collection: Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract levelTypes: bool option with get, set

        /// Represents the paragraph list item format.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] ListItemLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets or sets the level of the item in the list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract level: bool option with get, set
            /// Gets the list item bullet, number, or picture as a string. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listString: bool option with get, set
            /// Gets the list item order number in relation to its siblings. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract siblingIndex: bool option with get, set

        /// Represents a single paragraph in a selection, range, content control, or document body.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] ParagraphLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract list: Word.Interfaces.ListLoadOptions option with get, set
            /// Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItem: Word.Interfaces.ListItemLoadOptions option with get, set
            /// Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItemOrNullObject: Word.Interfaces.ListItemLoadOptions option with get, set
            /// Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listOrNullObject: Word.Interfaces.ListLoadOptions option with get, set
            /// Gets the parent body of the paragraph.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the table that contains the paragraph. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table that contains the paragraph. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract alignment: bool option with get, set
            /// Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
            /// 
            /// [Api set: WordApi 1.1]
            abstract firstLineIndent: bool option with get, set
            /// Indicates the paragraph is the last one inside its parent body. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isLastParagraph: bool option with get, set
            /// Checks whether the paragraph is a list item. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isListItem: bool option with get, set
            /// Gets or sets the left indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract leftIndent: bool option with get, set
            /// Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineSpacing: bool option with get, set
            /// Gets or sets the amount of spacing, in grid lines, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitAfter: bool option with get, set
            /// Gets or sets the amount of spacing, in grid lines, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitBefore: bool option with get, set
            /// Gets or sets the outline level for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract outlineLevel: bool option with get, set
            /// Gets or sets the right indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract rightIndent: bool option with get, set
            /// Gets or sets the spacing, in points, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceAfter: bool option with get, set
            /// Gets or sets the spacing, in points, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceBefore: bool option with get, set
            /// Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tableNestingLevel: bool option with get, set
            /// Gets the text of the paragraph. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set

        /// Contains a collection of {@link Word.Paragraph} objects.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] ParagraphCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract list: Word.Interfaces.ListLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItem: Word.Interfaces.ListItemLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listItemOrNullObject: Word.Interfaces.ListItemLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list.
            /// 
            /// [Api set: WordApi 1.3]
            abstract listOrNullObject: Word.Interfaces.ListLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent body of the paragraph.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the paragraph. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the paragraph. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
            /// 
            /// [Api set: WordApi 1.1]
            abstract alignment: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
            /// 
            /// [Api set: WordApi 1.1]
            abstract firstLineIndent: bool option with get, set
            /// For EACH ITEM in the collection: Indicates the paragraph is the last one inside its parent body. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isLastParagraph: bool option with get, set
            /// For EACH ITEM in the collection: Checks whether the paragraph is a list item. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isListItem: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the left indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract leftIndent: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineSpacing: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitAfter: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract lineUnitBefore: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the outline level for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract outlineLevel: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the right indent value, in points, for the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract rightIndent: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the spacing, in points, after the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceAfter: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the spacing, in points, before the paragraph.
            /// 
            /// [Api set: WordApi 1.1]
            abstract spaceBefore: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// For EACH ITEM in the collection: Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract tableNestingLevel: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text of the paragraph. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set

        /// Represents a contiguous area in a document.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] RangeLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets the parent body of the range.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the range. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the table that contains the range. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the table cell that contains the range. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table that contains the range. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.3]
            abstract hyperlink: bool option with get, set
            /// Checks whether the range length is zero. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isEmpty: bool option with get, set
            /// Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// Gets the text of the range. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set

        /// Contains a collection of {@link Word.Range} objects.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] RangeCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.1]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent body of the range.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.1]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the range. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the range. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the range. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains the range. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
            /// 
            /// [Api set: WordApi 1.3]
            abstract hyperlink: bool option with get, set
            /// For EACH ITEM in the collection: Checks whether the range length is zero. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isEmpty: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.1]
            abstract style: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// For EACH ITEM in the collection: Gets the text of the range. Read-only.
            /// 
            /// [Api set: WordApi 1.1]
            abstract text: bool option with get, set

        /// Specifies the options to be included in a search operation.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] SearchOptionsLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignorePunct: bool option with get, set
            /// Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract ignoreSpace: bool option with get, set
            /// Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchCase: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchPrefix: bool option with get, set
            /// Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchSuffix: bool option with get, set
            /// Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWholeWord: bool option with get, set
            /// Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
            /// 
            /// [Api set: WordApi 1.1]
            abstract matchWildcards: bool option with get, set

        /// Represents a section in a Word document.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] SectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the body object of the section. This does not include the header/footer and other section metadata.
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set

        /// Contains the collection of the document's {@link Word.Section} objects.
        /// 
        /// [Api set: WordApi 1.1]
        type [<AllowNullLiteral>] SectionCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the body object of the section. This does not include the header/footer and other section metadata.
            /// 
            /// [Api set: WordApi 1.1]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set

        /// Represents a table in a Word document.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets the parent body of the table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the content control that contains the table. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the content control that contains the table. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// Gets the table that contains this table. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the table cell that contains this table. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// Gets the table that contains this table. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract alignment: bool option with get, set
            /// Gets and sets the number of header rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract headerRowCount: bool option with get, set
            /// Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// Indicates whether all of the table rows are uniform. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isUniform: bool option with get, set
            /// Gets the nesting level of the table. Top-level tables have level 1. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract nestingLevel: bool option with get, set
            /// Gets the number of rows in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowCount: bool option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract style: bool option with get, set
            /// Gets and sets whether the table has banded columns.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedColumns: bool option with get, set
            /// Gets and sets whether the table has banded rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedRows: bool option with get, set
            /// Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// Gets and sets whether the table has a first column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleFirstColumn: bool option with get, set
            /// Gets and sets whether the table has a last column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleLastColumn: bool option with get, set
            /// Gets and sets whether the table has a total (last) row with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleTotalRow: bool option with get, set
            /// Gets and sets the text values in the table, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: bool option with get, set
            /// Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set
            /// Gets and sets the width of the table in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: bool option with get, set

        /// Contains the collection of the document's Table objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent body of the table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentBody: Word.Interfaces.BodyLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the table. Throws an error if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControl: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the content control that contains the table. Returns a null object if there isn't a parent content control.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentContentControlOrNullObject: Word.Interfaces.ContentControlLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains this table. Throws an error if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains this table. Throws an error if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCell: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableCellOrNullObject: Word.Interfaces.TableCellLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the table that contains this table. Returns a null object if it is not contained in a table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTableOrNullObject: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract alignment: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the number of header rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract headerRowCount: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// For EACH ITEM in the collection: Indicates whether all of the table rows are uniform. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isUniform: bool option with get, set
            /// For EACH ITEM in the collection: Gets the nesting level of the table. Top-level tables have level 1. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract nestingLevel: bool option with get, set
            /// For EACH ITEM in the collection: Gets the number of rows in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowCount: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract style: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets whether the table has banded columns.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedColumns: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets whether the table has banded rows.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBandedRows: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleBuiltIn: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets whether the table has a first column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleFirstColumn: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets whether the table has a last column with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleLastColumn: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets whether the table has a total (last) row with a special style.
            /// 
            /// [Api set: WordApi 1.3]
            abstract styleTotalRow: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the text values in the table, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the width of the table in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: bool option with get, set

        /// Represents a row in a Word document.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableRowLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// Gets parent table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellCount: bool option with get, set
            /// Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isHeader: bool option with get, set
            /// Gets and sets the preferred height of the row in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract preferredHeight: bool option with get, set
            /// Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: bool option with get, set
            /// Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// Gets and sets the text values in the row, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: bool option with get, set
            /// Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set

        /// Contains the collection of the document's TableRow objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableRowCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
            /// 
            /// [Api set: WordApi 1.3]
            abstract font: Word.Interfaces.FontLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets parent table.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the number of cells in the row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellCount: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// For EACH ITEM in the collection: Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
            /// 
            /// [Api set: WordApi 1.3]
            abstract isHeader: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the preferred height of the row in points.
            /// 
            /// [Api set: WordApi 1.3]
            abstract preferredHeight: bool option with get, set
            /// For EACH ITEM in the collection: Gets the index of the row in its parent table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the text values in the row, as a 2D Javascript array.
            /// 
            /// [Api set: WordApi 1.3]
            abstract values: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set

        /// Represents a table cell in a Word document.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableCellLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets the body object of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set
            /// Gets the parent row of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentRow: Word.Interfaces.TableRowLoadOptions option with get, set
            /// Gets the parent table of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellIndex: bool option with get, set
            /// Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
            /// 
            /// [Api set: WordApi 1.3]
            abstract columnWidth: bool option with get, set
            /// Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: bool option with get, set
            /// Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// Gets and sets the text of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: bool option with get, set
            /// Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set
            /// Gets the width of the cell in points. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: bool option with get, set

        /// Contains the collection of the document's TableCell objects.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableCellCollectionLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// For EACH ITEM in the collection: Gets the body object of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract body: Word.Interfaces.BodyLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent row of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentRow: Word.Interfaces.TableRowLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the parent table of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract parentTable: Word.Interfaces.TableLoadOptions option with get, set
            /// For EACH ITEM in the collection: Gets the index of the cell in its row. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract cellIndex: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
            /// 
            /// [Api set: WordApi 1.3]
            abstract columnWidth: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract horizontalAlignment: bool option with get, set
            /// For EACH ITEM in the collection: Gets the index of the cell's row in the table. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract rowIndex: bool option with get, set
            /// For EACH ITEM in the collection: Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
            /// 
            /// [Api set: WordApi 1.3]
            abstract shadingColor: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the text of the cell.
            /// 
            /// [Api set: WordApi 1.3]
            abstract value: bool option with get, set
            /// For EACH ITEM in the collection: Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
            /// 
            /// [Api set: WordApi 1.3]
            abstract verticalAlignment: bool option with get, set
            /// For EACH ITEM in the collection: Gets the width of the cell in points. Read-only.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: bool option with get, set

        /// Specifies the border style.
        /// 
        /// [Api set: WordApi 1.3]
        type [<AllowNullLiteral>] TableBorderLoadOptions =
            /// Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
            abstract ``$all``: bool option with get, set
            /// Gets or sets the table border color.
            /// 
            /// [Api set: WordApi 1.3]
            abstract color: bool option with get, set
            /// Gets or sets the type of the table border.
            /// 
            /// [Api set: WordApi 1.3]
            abstract ``type``: bool option with get, set
            /// Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
            /// 
            /// [Api set: WordApi 1.3]
            abstract width: bool option with get, set

    /// The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
    type [<AllowNullLiteral>] RequestContext =
        inherit OfficeCore.RequestContext
        abstract document: Document
        abstract application: Application

    /// The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
    type [<AllowNullLiteral>] RequestContextStatic =
        [<Emit "new $0($1...)">] abstract Create: ?url: string -> RequestContext

    type [<AllowNullLiteral>] BodySearch =
        abstract ignorePunct: bool option with get, set
        abstract ignoreSpace: bool option with get, set
        abstract matchCase: bool option with get, set
        abstract matchPrefix: bool option with get, set
        abstract matchSuffix: bool option with get, set
        abstract matchWholeWord: bool option with get, set
        abstract matchWildcards: bool option with get, set