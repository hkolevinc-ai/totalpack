# TotalPack → Temu scraper

Скрейпърът събира **само SKU кодовете от `totalpack_product_skus.xlsx`** и попълва приложения Temu шаблон. Не обхожда целия каталог на TotalPack.

## Стартиране в GitHub

1. Създай ново GitHub repository.
2. Качи **съдържанието** на тази папка в основната директория на repository-то. Папката `.github` също трябва да бъде качена.
3. Отвори **Actions → Run TotalPack Temu scraper → Run workflow**.
4. Остави `Product limit` на `0`, за да се обработят всички SKU.
5. След края на изпълнението отвори run-а и изтегли artifact **`totalpack-temu-results`**.

## Полета при ръчно стартиране

- `Product limit`: празно/`0` = всички SKU; число = само първите N.
- `Search`: незадължителен филтър по SKU, име или категория.
- `Exclude out-of-stock products`: ако е изключено, неналичните продукти остават с Quantity `0`.
- `Country/Region of Origin`: попълва задължителната Temu колона. Стойността трябва да бъде проверена за продуктите преди качване.

## Какво се попълва

- актуални име, описание, цена, наличност и изображения от WooCommerce Store API на TotalPack;
- тегло от подадения SKU файл;
- размери, брой в пакет и материал от продуктовия текст, когато са налични;
- Temu категория според категорията на продукта в TotalPack;
- `boxnow`, `ULTRAPAK`, стандартен данъчен код и останалите технически полета, налични в подадения шаблон;
- EU Responsible person остава празно, защото трябва да съвпада с лице, предварително регистрирано в Seller Center.

Когато сайтът не съдържа пакетен размер/дебелина, скрейпърът използва техническа резервна стойност и я описва в `totalpack_scrape_report.csv`. Проверявай редовете със статус `OK_WITH_FALLBACKS` преди качване в Temu.

Версия 1.1 проверява автоматично допустимите Temu стойности за калцуни, маски и стреч фолио и не позволява текстът на линкове към свързани категории да промени материала на продукта.

## Резултати

- `TEMU_TOTALPACK_UPLOAD_part_001.xlsx` — файл за качване в Temu;
- `totalpack_scrape_report.csv` — статус за всеки поискан SKU;
- `summary.json` — кратка статистика от изпълнението.

Файловете се разделят автоматично при повече от 1900 реда, като всички редове на един продукт остават в една част.

## Локално стартиране

```bash
python -m pip install -r requirements.txt
python totalpack_scraper.py \
  --template TEMU_ALL-PURPOSE-LABELS_ALUMINUM-FOIL_BOXES_BOXES_CLOTH-FACE-MASKS_CONTAINERS_CONTAINERS_DISPOSABLE-MEDIC.xlsx \
  --sku-file totalpack_product_skus.xlsx \
  --output-dir output
```
