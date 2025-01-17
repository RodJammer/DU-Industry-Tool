## Mercury edition
- Versions 0.500+ by tobitege (https://github.com/tobitege/DU-Industry-Tool)
- Major updates to recipes to be current with patch 0.31.x "Mercury".
- Schematic cost is now maintained and associated with the corresponding pures, products and elements.
- Calculated cost output now also lists schematic costs per tier and category.
- Display of a Total Ingredients List.
- Ingredients as links for "drill down".
- Listing of all "superior" recipes which have the item as an ingredient ("drill up").
- Display of mass, volume (of item itself) and nanocraftable status.
- Results can be copied to clipboard.
- Revamped UI with tabs support and auto-completion search box.

Big thanks to Jericho1060 for making available item/recipe data dumps via
his website https://du-lua.dev which helped me a lot to update the data displayed!

Binary releases available here (v0.500+):
https://github.com/tobitege/DU-Industry-Tool/releases

![Mercury Overview](https://github.com/tobitege/DU-Industry-Tool/blob/master/DU%20Industry%20Tool%20(Mercury)%20Screenshot.png?raw=true)

# DU-Industry-Tool
Basic WinForms project that allows you to view calculated values of all known in-game recipes. Has search bar, and ability to drill down through intermediates to find their recipes as well

GIF showing functionality
https://gyazo.com/a8740425ac2fe244d87e87980c16a2cf


# Older releases < 0.5.x are no longer maintained

Anyone could add the mining units or other new items, it's just a json file - but please share it, either through a fork or a PR.  Either add the new stuff manually, or:
1. Do a hyperion export of all item NqRecipeId's for all items (particularly the new ones), along with Name
2. Create a lua script to run core.getSchematicInfo(id) for every NqRecipeId.  Store all the data it gives you, and export it to json when done (writing to logfile is usually best to be able to copy it out)
3. Do a hyperion export of all item Groups (it's a separate category in lua export), the GroupId and ParentGroupId (see: Groups.json in this repo, but there may be new ones)
4. Using any language, lua or out of game at this point, load the recipe info json file, along with the Groups json file, and combine them.  Add fields as necessary to the recipe data to match the 'SchematicRecipe' class in Recipe.cs, such as 'Name', 'GroupId', and 'ParentGroupName' (Key doesn't need to exist and/or should be null).  I think those three vars are the only thing different from core.getSchematicInfo's output
5. Save as RecipesGroups.json and everything should work
