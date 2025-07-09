let
    // Base source pulling from the wiki API for items
    BaseUrl = "https://ringofbrodgar.com/api.php?action=query&list=categorymembers&cmtitle=Category:Objects&cmlimit=500&format=json",
    GetPage = (cmContinue as text) =>
        let
            Url = if cmContinue = "" then BaseUrl else BaseUrl & "&cmcontinue=" & cmContinue,
            Source = Json.Document(Web.Contents(Url)),
            Data = Source[query][categorymembers],
            NextContinue = try Source[continue][cmcontinue] otherwise null
        in
            [Data=Data, NextContinue=NextContinue],

    Pages = List.Generate(
        ()=> GetPage(""),
        each _[NextContinue] <> null,
        each GetPage(_[NextContinue]),
        each _[Data]
    ),

    AllData = List.Combine(Pages),
    Items = Table.FromRecords(AllData),
    Cleaned = Table.RemoveColumns(Items,{"pageid", "ns"}),
    WithURL = Table.AddColumn(Cleaned, "URL", each "https://ringofbrodgar.com/wiki/" & Text.Replace([title], " ", "_")),

    // Pull the first 100 entries for demo (can be made dynamic)
    Batch = Table.Range(WithURL, 0, 3000),

    // Pull HTML content and parse tables
    WithPage = Table.AddColumn(Batch, "Page", each try Web.Page(Web.Contents([URL])) otherwise null),
    ExpandPage = Table.ExpandTableColumn(WithPage, "Page", {"Data"}, {"Page.Data"}),
    ExpandData = Table.ExpandTableColumn(ExpandPage, "Page.Data", {"Column1", "Column2", "Column3"}, {"Col1", "Col2", "Col3"}),

    // Fields to extract
    FieldsToExtract = {
        "Object(s) Required", "Produced By",
        "Base LP", "Mental Weight", "EXP Cost", "LP / Hour (real)",
        "Energy Filled", "Hunger Filled", "Satiates", "Gilding Chance",
        "Gilding Attributes", "Gilding Bonus 1", "Gilding Bonus 2",
        "Gilding Bonus 3", "Gilding Bonus 4"
    },

    Filtered = Table.SelectRows(ExpandData, each 
        List.Contains({"Object(s) Required", "Produced By"}, [Col2]) or 
        List.Contains(FieldsToExtract, [Col1])
    ),

    // Normalize label and value
    AddLabel = Table.AddColumn(Filtered, "Label", each if List.Contains({"Object(s) Required", "Produced By"}, [Col2]) then [Col2] else [Col1]),
    AddValue = Table.AddColumn(AddLabel, "Value", each if [Label] = [Col2] then [Col3] else [Col2]),
    KeepRelevant = Table.SelectColumns(AddValue, {"title", "Label", "Value"}),

    // Pivot labels to columns
    Pivoted = Table.Pivot(KeepRelevant, List.Distinct(KeepRelevant[Label]), "Label", "Value"),

    // Clean comma/quantity formatting
    CleanText = (input as nullable text) as nullable text =>
        let
            removed = List.Accumulate(
                {" x12", " x11", " x10", " x9", " x8", " x7", " x6", " x5", " x4", " x3", " x2", " x1"},
                input,
                (state, qty) => if state <> null then Text.Replace(state, qty, "") else null
            ),
            newline = if removed <> null then Text.Replace(removed, ", ", "#(lf)") else null
        in
            Text.Trim(newline),

    WithCleaned = Table.TransformColumns(Pivoted, {
    {"Object(s) Required", each CleanText(_), type text},
    {"Produced By", each CleanText(_), type text}
}),

// Rename 'title' to 'Name' in the first table before join
WithURLRenamed = Table.RenameColumns(WithURL, {{"title", "Name"}}),

// Join on the renamed column 'Name' and original 'title'
Joined = Table.Join(WithURLRenamed, "Name", WithCleaned, "title", JoinKind.LeftOuter)

in
    Joined
