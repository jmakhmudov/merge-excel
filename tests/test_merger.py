from datetime import datetime

from merge_excel.merger import export
from merge_excel.profile import Profile, ResultColumn
from merge_excel.scanner import scan

from conftest import link, read, template_with


def noop(*_):
    pass


def test_continuation_rows_are_joined_and_footer_is_removed(source, tmp_path):
    folder = source(
        "service.xlsx",
        {
            "Main": [
                ["Order", "Customer", "Work", "Cost"],
                ["RO-1", "Alice", "Oil change", 100],
                [None, None, "Oil filter", 20.0],
                [None, "ignored", "Engine oil", 50],
                ["RO-2", "Bob", "Inspection", 70],
                ["Total", None, None, 240],
            ]
        },
    )
    result = scan(folder, noop)
    profile = Profile(
        columns=[ResultColumn("Order"), ResultColumn("Customer"), ResultColumn("Work", join=True), ResultColumn("Cost", join=True, separator="; ")],
        skip=[],
    )
    link(profile, result, result.templates[0], {"0": "Order", "1": "Customer", "2": "Work", "3": "Cost"})
    outputs = export(result, profile, tmp_path / "out.xlsx", noop)
    assert [output.rows for output in outputs] == [2]
    assert read(outputs[0].path) == [
        ["Order", "Customer", "Work", "Cost"],
        ["RO-1", "Alice", "Oil change\nOil filter\nEngine oil", "100; 20; 50"],
        ["RO-2", "Bob", "Inspection", "70"],
    ]


def test_skip_prefixes_multiple_sources_and_sheet_name(source, tmp_path):
    folder = source(
        "parts.xlsx",
        {
            "North": [
                ["Order", "Labor", "Part", "Date"],
                ["RO-1", "Oil change", None, datetime(2024, 1, 2)],
                ["Repeat repairs: none", None, None, None],
                [None, None, "04152-YZZA1", None],
                ["RO-2", "Wash", None, datetime(2024, 1, 3)],
            ]
        },
    )
    result = scan(folder, noop)
    profile = Profile(columns=[ResultColumn("Order"), ResultColumn("Lines", join=True, separator=", "), ResultColumn("Branch")], skip=["repeat"])
    link(profile, result, result.templates[0], {"0": "Order", "1": "Lines", "2": "Lines", "sheet": "Branch"})
    rows = read(export(result, profile, tmp_path / "out.xlsx", noop)[0].path)
    assert rows[1:] == [["RO-1", "Oil change, 04152-YZZA1", "North"], ["RO-2", "Wash", "North"]]


def test_unique_values_and_latest_record(source, tmp_path):
    folder = source(
        "service.xlsx",
        {
            "Main": [
                ["Order", "VIN", "Opened"],
                ["RO-1", "VIN-A", "03.06.2024 11:23:10"],
                ["RO-2", "VIN-B", 45839],
                ["RO-3", "VIN-A", "2025-01-05"],
                ["RO-4", "VIN-A", None],
                ["RO-5", "VIN-C", "garbage"],
                ["RO-6", "VIN-C", "01.01.2024"],
                ["RO-7", None, "01.01.2026"],
                ["RO-8", "VIN-D", None],
                ["RO-9", "-", "01.01.2026"],
            ]
        },
    )
    result = scan(folder, noop)
    profile = Profile(columns=[ResultColumn("Order"), ResultColumn("VIN"), ResultColumn("Opened", date=True)], skip=[])
    link(profile, result, result.templates[0], {"0": "Order", "1": "VIN", "2": "Opened"})
    profile.unique, profile.latest, profile.latest_by = profile.columns[1].id, True, profile.columns[2].id
    merged, unique, latest = export(result, profile, tmp_path / "Merged.xlsx", noop)
    assert (merged.rows, unique.rows, latest.rows) == (9, 4, 4)
    assert (unique.path.name, latest.path.name) == ("Merged - unique VIN.xlsx", "Merged - latest by Opened.xlsx")
    assert read(unique.path) == [["VIN-A"], ["VIN-B"], ["VIN-C"], ["VIN-D"]]
    assert read(latest.path) == [
        ["Order", "VIN", "Opened"],
        ["RO-2", "VIN-B", datetime(2025, 7, 1)],
        ["RO-3", "VIN-A", datetime(2025, 1, 5)],
        ["RO-6", "VIN-C", datetime(2024, 1, 1)],
        ["RO-8", "VIN-D", None],
    ]


def test_extract_and_date_options(source, tmp_path):
    folder = source(
        "orders.xlsx",
        {"Main": [["Order", "Note", "Body"], ["Repair order RO-000101", "opened 03.01.2023 10:12:41 by SA", "VIN 4T1BF3EK8BR105031"]]},
    )
    result = scan(folder, noop)
    profile = Profile(
        columns=[
            ResultColumn("Order", extract="regex", pattern=r"RO-\d+"),
            ResultColumn("Opened", extract="regex", pattern=r"opened (\S+ \S+)", date=True),
            ResultColumn("VIN", extract="last", pattern="17"),
        ],
        skip=[],
    )
    link(profile, result, result.templates[0], {"0": "Order", "1": "Opened", "2": "VIN"})
    rows = read(export(result, profile, tmp_path / "out.xlsx", noop)[0].path)
    assert rows[1] == ["RO-000101", datetime(2023, 1, 3, 10, 12, 41), "4T1BF3EK8BR105031"]


def test_auto_link_by_header_name_skips_ambiguous_names(source):
    source("a.xlsx", {"Main": [["Order", "Customer", "Total"], ["RO-1", "Alice", 10]]})
    folder = source("b.xlsx", {"Main": [["Order", "Customer", "Total", "Total", "Extra"], ["RO-2", "Bob", 1, 2, "x"]]})
    result = scan(folder, noop)
    profile = Profile(columns=[ResultColumn("Order"), ResultColumn("Customer"), ResultColumn("Total")])
    link(profile, result, template_with(result, "Order"), {"0": "Order", "1": "Customer", "2": "Total"})
    other = template_with(result, "Extra")
    ids = {column.name: column.id for column in profile.columns}
    assert profile.is_auto(other)
    assert profile.links_for(other, result.templates) == {"0": ids["Order"], "1": ids["Customer"]}
