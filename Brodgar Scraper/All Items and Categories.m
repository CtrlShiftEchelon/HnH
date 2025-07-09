let
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
    Result = Table.FromRecords(AllData),
    #"Removed Columns" = Table.RemoveColumns(Result,{"pageid", "ns"}),
    // Build URL from title
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "URL", each "https://ringofbrodgar.com/wiki/" & Text.Replace([title], " ", "_")),

    // Extract categories using your function
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Categories", each GetCategories([URL])),

    // Filter out unwanted categories
    #"Filtered Categories" = Table.TransformColumns(#"Added Custom1", {
        {"Categories", each
            let
                cats = Text.Split(_, "#(lf)"),
                removed = List.Select(cats, each 
                    _ <> "Objects" and 
                    _ <> "Stockpile" and 
                    _ <> "Cheeses" and 
                    _ <> "GenericTypePage" and 
                    _ <> "Article_stubs" and 
                    not Text.EndsWith(_, "_Structures")
                ),
                result = Text.Combine(removed, "#(lf)")
            in
                result
        }
    })
in
    #"Filtered Categories"
