let
    Source = Excel.CurrentWorkbook(){[Name="Table10"]}[Content],
    #"Added Custom" = Table.AddColumn(Source, "Page", each Web.Page(Web.Contents([URL]))),
    #"Expanded Page" = Table.ExpandTableColumn(#"Added Custom", "Page", {"Data"}, {"Page.Data"}),
    #"Expanded Page.Data" = Table.ExpandTableColumn(#"Expanded Page", "Page.Data", {"Column1", "Column2", "Column3"}, {"Page.Data.Column1", "Page.Data.Column2", "Page.Data.Column3"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Page.Data", each 
        ([Page.Data.Column2] <> "Discovery Req.This(<i>discoveryreq</i>) is only for items that need to be discovered in addition to items listed in ""Object(s) Required""/<i>(<i>objectsreq</i>).</i><small><br><b><i>(Temporary active on all pages, but leave empty if ""None"".</i></b><small>" and 
         [Page.Data.Column2] <> "QL10 Equipment Statistics" and 
         [Page.Data.Column2] <> "Required By" and 
         [Page.Data.Column2] <> "Size (item)Item (inventory)size.<br><br>xitem & yitem" and 
         [Page.Data.Column2] <> "Skill(s) RequiredSpecific needed skills that enable a given object or item." and 
         [Page.Data.Column2] <> "Specific Type of" and 
         [Page.Data.Column2] <> "Stockpile Stockpile data for (all) items is set at the Stockpile page in the Stockpile property data section.<br><br><i>Please use ""Source Edit"" mode when editing there.</i>") and 
         ([Page.Data.Column1] = "Armor Penetration" or [Page.Data.Column1] = "Damage" or [Page.Data.Column1] = "Range" or [Page.Data.Column1] = "Vital statistics")
    ),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows", "Custom", each if [Page.Data.Column1] = "Vital statistics" and [Page.Data.Column3] <> null then [Page.Data.Column2] else [Page.Data.Column1]),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom1",{{"Custom", "FinalLabel"}}),
    #"Added IsTableFlag" = Table.AddColumn(#"Renamed Columns", "IsTableFlag", each if Value.Is([FinalLabel], type table) then "IsTable" else "NotTable"),
    #"Filtered Out Tables" = Table.SelectRows(#"Added IsTableFlag", each ([IsTableFlag] = "NotTable")),
    #"Removed IsTableFlag" = Table.RemoveColumns(#"Filtered Out Tables",{"IsTableFlag"}),
    #"Added Custom2" = Table.AddColumn(#"Removed IsTableFlag", "Custom", each if [Page.Data.Column1] = "Vital statistics" and [Page.Data.Column3] <> null then [Page.Data.Column3] else [Page.Data.Column2]),
    #"Renamed Columns1" = Table.RenameColumns(#"Added Custom2",{{"Custom", "FinalValue"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"Page.Data.Column1", "Page.Data.Column2", "Page.Data.Column3"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"FinalLabel", type text}}),

    #"Pivoted Column" = Table.Pivot(#"Changed Type", List.Distinct(#"Changed Type"[FinalLabel]), "FinalLabel", "FinalValue"),

    // Rebuild "Item Requirements" dynamically combining "Object(s) Required", "Produced By", plus special cases
    #"Added Custom3" = Table.AddColumn(#"Pivoted Column", "Item Requirements", each 
        let
            objReq = if [#"Object(s) Required"] = null then "" else [#"Object(s) Required"],
            prodBy = if [#"Produced By"] = null then "" else [#"Produced By"],
            addAnvil = if Text.Contains(prodBy, "Anvil") then "Anvil + Hammer" else "",
            addKiln = if Text.Contains(prodBy, "Kiln") then "Kiln + Fuel" else "",
            additionsList = List.RemoveItems({addAnvil, addKiln}, {""}),
            combinedAdditions = if List.Count(additionsList) = 0 then "" else Text.Combine(additionsList, ", "),
            finalList = List.RemoveItems({combinedAdditions, objReq}, {""})
        in
            if List.Count(finalList) = 0 then "" else Text.Combine(finalList, ", ")
    ),

    // Remove original columns now that "Item Requirements" is rebuilt
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom3",{"Object(s) Required", "Produced By"}),

    // Clean "Item Requirements": remove " x1"..." x12" and replace commas with line breaks
    #"Cleaned Item Requirements" = Table.TransformColumns(
        #"Removed Columns1",
        {
            {"Item Requirements", each 
                let
                    removedMultipliers = List.Accumulate(
                        {" x12", " x11", " x10", " x9", " x8", " x7", " x6", " x5", " x4", " x3", " x2", " x1"},
                        _,
                        (state, current) => Text.Replace(state, current, "")
                    ),
                    replacedLineBreaks = Text.Replace(removedMultipliers, ", ", "#(lf)")
                in
                    Text.Trim(replacedLineBreaks)
            }
        }
    ),

    // Split "Item Requirements" into Q of Ing1 ... Q of Ing5
    #"Added Q of Ing1" = Table.AddColumn(#"Cleaned Item Requirements", "Q of Ing1", each try Text.Split([Item Requirements], "#(lf)"){0} otherwise null),
    #"Added Q of Ing2" = Table.AddColumn(#"Added Q of Ing1", "Q of Ing2", each try Text.Split([Item Requirements], "#(lf)"){1} otherwise null),
    #"Added Q of Ing3" = Table.AddColumn(#"Added Q of Ing2", "Q of Ing3", each try Text.Split([Item Requirements], "#(lf)"){2} otherwise null),
    #"Added Q of Ing4" = Table.AddColumn(#"Added Q of Ing3", "Q of Ing4", each try Text.Split([Item Requirements], "#(lf)"){3} otherwise null),
    #"Added Q of Ing5" = Table.AddColumn(#"Added Q of Ing4", "Q of Ing5", each try Text.Split([Item Requirements], "#(lf)"){4} otherwise null),

    // Rename Damage to baseDamage
    #"Renamed Damage" = Table.RenameColumns(#"Added Q of Ing5", {{"Damage", "baseDamage"}}),

    // Add Quality column as 2nd column, all values = 10
    #"Added Quality" = Table.AddColumn(#"Renamed Damage", "Quality", each 10, Int64.Type),

    // Add Damage (calculated) as 3rd column using formula: Damage = baseDamage * sqrt(sqrt(10 * Quality)/10)
    #"Added Damage" = Table.AddColumn(#"Added Quality", "Damage", each 
        let
            baseDmgNum = try Number.FromText([baseDamage]) otherwise 0,
            qualityNum = try Number.FromText(Text.From([Quality])) otherwise 10,
            calc = baseDmgNum * Number.Sqrt( Number.Sqrt(10 * qualityNum) / 10 )
        in
            calc, type number),

    // Add Potential Quality column after Q of Ing5, all values = 0
    #"Added Potential Quality" = Table.AddColumn(#"Added Damage", "Potential Quality", each 0, Int64.Type),

    // Reorder columns as requested
    #"Reordered Columns" = Table.ReorderColumns(#"Added Potential Quality", 
    {
        "Name", "URL", "Quality", "Damage", "Item Requirements", 
        "Q of Ing1", "Q of Ing2", "Q of Ing3", "Q of Ing4", "Q of Ing5", 
        "Potential Quality", "baseDamage"
    } 
    & List.RemoveItems(
        Table.ColumnNames(#"Added Potential Quality"), 
        {
            "Name", "URL", "Quality", "Damage", "Item Requirements", 
            "Q of Ing1", "Q of Ing2", "Q of Ing3", "Q of Ing4", "Q of Ing5", 
            "Potential Quality", "baseDamage"
        }
    )
)
in
    #"Reordered Columns"
