from extractors.fixture_extractor import FixtureExtractor


def test_provider_balance_fixtures_omit_blank_tail_and_identify_country_rows():
    capture = FixtureExtractor._balance_table(
        [
            {"col_a": "CDF", "col_b": None, "col_c": None},
            {"col_a": "Provider A", "col_b": None, "col_c": None},
            {"col_a": None, "col_b": None, "col_c": None},
        ],
        provider=True,
    )

    assert capture.row_count == 2
    assert capture.rows[0]["row_class"] == "country-row"
    assert capture.rows[1]["row_class"] == ""
    assert capture.html.count("<tr") == 2


def test_merchant_balance_fixture_identifies_country_rows():
    capture = FixtureExtractor._balance_table(
        [
            {"col_a": "Zambia", "col_b": "ZMW balance"},
            {"col_a": "Merchant A", "col_b": None},
        ],
        provider=False,
    )

    assert [row["row_class"] for row in capture.rows] == [
        "country-row",
        "",
    ]
