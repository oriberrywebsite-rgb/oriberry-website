# Oriberry menu and blog sheet schema

Production spreadsheet: `Oriberry_CMS_Database` (`1tCAB0w3bmQjOkSxkE-_U_42LAJaBpKmiOLHFaZTdSns`).
Backup made before the 2026-10-22 migration: `1hFhuvHf0e2KX9ELuY5fxv4oxHLN5piqKm2c1Ge8cjBA`.
Backup made before publishing the first blog articles on 2026-10-23: `1EQhW2Q8uZ8eUsu7z8lmu96CKidNW_c6UYDQkigTnIA0`.

## Product sizes

`Food_Drinks_Menu.normal_price` and `large_price` are the source of truth. A blank cell means that size is unavailable; numeric zero is a valid price. `price` is used only for legacy rows where **both** size cells are blank. Store override columns apply only to available sizes. `side_dish1–3` remain product recommendations.

## Addons

`Addons_Menu.normal_price` and `large_price` were added without deleting `category` or `price`. The old numeric `price` was copied into both size prices for the five existing addons. Blank means unavailable for that size; zero means free. Enter real numeric cells, not formatted text.

`Product_Addons` has `Product_ID`, `Addon_ID`, `Status`, `Mapping_Note`. Only `Active` rows permit an addon for a product. The initial 25 links between five Coffee products and five Coffee addons were inferred from the old `coffee` category and are **not a confirmed product-specific business configuration**. Edit or mark links `Inactive` to set each product's actual addons. Products without links have no addons. `side_dish1–3` are independent.

## Blog

The three homepage links map to existing CMS tabs as follows: The Business → `Block3_Shops_Menu`, The Farmers → `Block1_Farmers_Philosophy`, The Processing → `Block2_Roastery_Products`. This mapping is based on the tab topics and homepage labels; no previous link mapping existed.

The tabs retain their old columns. `Block1_Farmers_Philosophy` gained `Slug`, `Published_At`, `Sort_Order`; the other two gained those plus `Excerpt` and `Main_Content`. To publish a row, fill a unique lowercase hyphenated `Slug` in its tab, `Title`, `Excerpt`, `Main_Content`, an HTTPS `Cover_Image_URL`, and set `Status` to `Active` or `Published`. Empty cells and placeholder `test` content do not publish. `Sort_Order` is ascending when numeric; otherwise `Published_At` sorts descending. A slug is unique within a category because article URLs include both category and slug.

`Main_Content` accepts plain paragraphs, `##`/`###` headings, `- ` list items, and image lines like `![Alt text](https://example.com/photo.jpg)`. Raw HTML is shown as text, not executed. Each blog page displays three posts. Existing placeholder rows were preserved and are not public until completed.

On 2026-10-23, one introductory article was published in each category in row 3 (ID 2): `from-farm-to-shops`, `working-with-small-farmers`, and `exploring-processing-and-roasting`. Their copy and images were adapted from the existing public homepage; they do not represent new interviews or reporting. The owner can edit these CMS rows directly as fuller stories become available.

## Order fields

`tableNumber` is required for new DINE_IN requests. It is saved into the existing `Oriberry_DonHang_Store.Location_Detail` column. Delivery uses `locationDetail`; PICKUP uses neither. The `Items_Detail` text stores each line's size, addons, item note, quantity and line total. `Note` continues to store the order-level note. Old rows remain readable.
