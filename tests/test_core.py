from totalpack_scraper import SourceSku, category_properties, clean_html, parse_dimensions, parse_pack_count, pick_category


def test_clean_html_and_pack_count():
    text = clean_html("<p>Кутия за храна</p><p>50 броя в стек</p>")
    assert text == "Кутия за храна\n50 броя в стек"
    assert parse_pack_count(text) == 50


def test_dimensions_mm():
    source = SourceSku("1")
    assert parse_dimensions("Размер: 190/130/73мм", source, 10)[:3] == (19, 13, 7.3)


def test_category_mapping():
    product = {"name": "Касова ролка", "categories": [{"slug": "receipts"}]}
    assert pick_category(product, SourceSku("1")) == 9030


def test_mask_and_stretch_values_match_template_options():
    mask = {"name": "Еднократна маска", "categories": [{"slug": "maski-za-litse"}]}
    assert pick_category(mask, SourceSku("1")) == 54386
    assert category_properties(31121, "полиетилен", 20)["t_3_Property:12"] == "PE"
    assert category_properties(17189, "pvc stretch film", 20)["t_3_Property:121"] == "Polyvinyl Chloride Pvc"
    assert category_properties(17189, "lldpe stretch film", 20)["t_3_Property:121"] == "PE"
