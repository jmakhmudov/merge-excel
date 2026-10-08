from merge_excel.scanner import scan

from conftest import template_with


def noop(*_):
    pass


def test_group_row_under_title_and_description_row(source):
    folder = source(
        "report.xlsx",
        {
            "Data": [
                ["Branch report"],
                [None, "Client", None, "Work", None],
                ["Report column", "Name", "Category", "Description", "Cost"],
                ["Definition", "Full name", "private / fleet", "as in FRM", "dealer currency"],
                ["RO-1", "Alice", "PRIVATE", "Oil change", 100],
                [None, None, None, "Oil filter", 20],
                ["RO-2", "Bob", "FLEET", "Inspection", 70],
            ]
        },
        {"Data": ["B2:C2", "D2:E2"]},
    )
    template = scan(folder, noop).templates[0]
    header = template.sheets[0].header
    assert (header.start, header.end, header.body) == (1, 2, 4)
    assert template.names == ("Report column", "Client / Name", "Client / Category", "Work / Description", "Work / Cost")


def test_two_level_headers_with_and_without_vertical_merges_share_a_template(source):
    source(
        "plain.xlsx",
        {"Sheet": [["Name", "Cost", None, "Date"], [None, "Labor", "Parts", None], ["X1", 10, None, 45000]]},
        {"Sheet": ["B1:C1"]},
    )
    folder = source(
        "shifted.xlsx",
        {"Sheet": [[None, "Name", "Cost", None, "Date"], [None, None, "Labor", "Parts", None], [None, "Y1", 1, 2, 45001]]},
        {"Sheet": ["B1:B2", "C1:D1", "E1:E2"]},
    )
    result = scan(folder, noop)
    assert len(result.templates) == 1
    template = result.templates[0]
    assert template.names == ("Name", "Cost / Labor", "Cost / Parts", "Date")
    assert sorted(sheet.header.left for sheet in template.sheets) == [0, 1]


def test_banner_rows_before_the_first_record_are_skipped(source):
    folder = source(
        "banner.xlsx",
        {"Sheet": [["Order", "Item", "Qty"], ["Repeat repairs: none"], ["RO-1", "Bolt", 2], ["RO-2", "Nut", 4]]},
        {"Sheet": ["A2:C2"]},
    )
    header = scan(folder, noop).templates[0].sheets[0].header
    assert header.body == 2


def test_first_record_of_a_text_only_table_is_kept(source):
    folder = source("people.xlsx", {"Sheet": [["Person", "City", "Country"], ["Alice", "Paris", "France"], ["Bob", "Berlin", "Germany"]]})
    header = scan(folder, noop).templates[0].sheets[0].header
    assert header.body == 1


def test_key_only_first_record_is_kept(source):
    folder = source("grouped.xlsx", {"Sheet": [["Order", "Item", "Qty"], ["ORD-1"], [None, "Bolt", 2], ["ORD-2"], [None, "Nut", 4]]})
    header = scan(folder, noop).templates[0].sheets[0].header
    assert header.body == 1


def test_templates_and_ignored_sheets(source):
    source("a.xlsx", {"Main": [["Order", "Customer", "Total"], ["RO-1", "Alice", 10]], "Empty": []})
    source("b.xlsx", {"Main": [["ORDER", "CUSTOMER", "TOTAL"], ["RO-2", "Bob", 20]], "Notes": [["VIN-1"], ["VIN-2"]]})
    folder = source("c.xlsx", {"Main": [["Repair order", "Client", "Sum"], ["RO-3", "Eve", 30]]})
    result = scan(folder, noop)
    assert sorted(len(template.sheets) for template in result.templates) == [1, 2]
    assert {sheet.path.name for sheet in template_with(result, "Order").sheets} == {"a.xlsx", "b.xlsx"}
    assert sorted((item.path.name, item.sheet, item.reason) for item in result.ignored) == [
        ("a.xlsx", "Empty", "empty"),
        ("b.xlsx", "Notes", "no header found"),
    ]


def test_unnamed_data_columns_are_offered_for_linking(source):
    folder = source("tail.xlsx", {"Sheet": [["Order", "Customer", "Model"], ["RO-1", "Alice", "Camry", 30680, 1, 276]]})
    columns = scan(folder, noop).templates[0].columns(())
    assert [(column.letter, column.name) for column in columns if column.letter] == [
        ("A", "Order"), ("B", "Customer"), ("C", "Model"), ("D", ""), ("E", ""), ("F", ""),
    ]
