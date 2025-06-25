let
    Source = Excel.CurrentWorkbook(){[Name="Table10"]}[Content],
    #"Added Custom" = Table.AddColumn(Source, "Page", each Web.Page(Web.Contents([URL]))),
    #"Expanded Page" = Table.ExpandTableColumn(#"Added Custom", "Page", {"Data"}, {"Page.Data"}),
    #"Expanded Page.Data" = Table.ExpandTableColumn(#"Expanded Page", "Page.Data", {"Column1", "Column2", "Column3"}, {"Page.Data.Column1", "Page.Data.Column2", "Page.Data.Column3"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Page.Data", each ([Page.Data.Column2] <> "Discovery Req.This(<i>discoveryreq</i>) is only for items that need to be discovered in addition to items listed in ""Object(s) Required""/<i>(<i>objectsreq</i>).</i><small><br><b><i>(Temporary active on all pages, but leave empty if ""None"".</i></b><small>" and [Page.Data.Column2] <> "QL10 Equipment Statistics" and [Page.Data.Column2] <> "Required By" and [Page.Data.Column2] <> "Size (item)Item (inventory)size.<br><br>xitem & yitem" and [Page.Data.Column2] <> "Skill(s) RequiredSpecific needed skills that enable a given object or item." and [Page.Data.Column2] <> "Specific Type of" and [Page.Data.Column2] <> "Stockpile Stockpile data for (all) items is set at the Stockpile page in the Stockpile property data section.<br><br><i>Please use ""Source Edit"" mode when editing there.</i>") and ([Page.Data.Column1] = "Armor Penetration" or [Page.Data.Column1] = "Damage" or [Page.Data.Column1] = "Range" or [Page.Data.Column1] = "Vital statistics")),
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
    #"Added Custom3" = Table.AddColumn(#"Pivoted Column", "Item Requirements", each let
    objReq = if [#"Object(s) Required"] = null then "" else [#"Object(s) Required"],
    prodBy = if [#"Produced By"] = null then "" else [#"Produced By"],
    addAnvil = if Text.Contains(prodBy, "Anvil") then "Anvil + Hammer" else "",
    addKiln = if Text.Contains(prodBy, "Kiln") then "Kiln + Fuel" else "",
    additionsList = List.RemoveItems({addAnvil, addKiln}, {""}),
    combinedAdditions = if List.Count(additionsList) = 0 then "" else Text.Combine(additionsList, ", "),
    finalList = List.RemoveItems({combinedAdditions, objReq}, {""})
in
    if List.Count(finalList) = 0 then "" else Text.Combine(finalList, ", ")),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom3",{"Name", "URL", "Item Requirements", "Object(s) Required", "Produced By", "Slot(s) Occupied", "Damage", "Armor Penetration", "Range"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns",{"Object(s) Required", "Produced By"})
in
    #"Removed Columns1"
