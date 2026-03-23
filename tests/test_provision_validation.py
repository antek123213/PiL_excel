import export_automaty


def test_validate_provision_sql_vs_xls_success(monkeypatch):
    def fake_sql(*args, **kwargs):
        return {
            1101: {"prowizja_zl": 10.0},
            1102: {"prowizja_zl": 20.0},
        }

    def fake_xls(_):
        return {
            1101: 10.0,
            1102: 20.0,
        }

    monkeypatch.setattr(export_automaty, "get_monthly_commission", fake_sql)
    monkeypatch.setattr(export_automaty, "load_reference_provision_from_xls", fake_xls)

    comparison, exit_code = export_automaty.validate_provision_sql_vs_xls(
        conn=None,
        month_str="2026-02",
        commission_source={},
        xls_path="dummy.xls",
        strict_month_window=True,
    )

    assert exit_code == 0
    assert comparison["missing_in_sql"] == []
    assert comparison["extra_in_sql"] == []
    assert comparison["mismatches"] == []


def test_validate_provision_sql_vs_xls_detects_differences(monkeypatch):
    def fake_sql(*args, **kwargs):
        return {
            1101: {"prowizja_zl": 10.0},
            1103: {"prowizja_zl": 30.0},
        }

    def fake_xls(_):
        return {
            1101: 11.0,
            1102: 20.0,
        }

    monkeypatch.setattr(export_automaty, "get_monthly_commission", fake_sql)
    monkeypatch.setattr(export_automaty, "load_reference_provision_from_xls", fake_xls)

    comparison, exit_code = export_automaty.validate_provision_sql_vs_xls(
        conn=None,
        month_str="2026-02",
        commission_source={},
        xls_path="dummy.xls",
        strict_month_window=True,
    )

    assert exit_code == 1
    assert 1102 in comparison["missing_in_sql"]
    assert 1103 in comparison["extra_in_sql"]
    assert len(comparison["mismatches"]) == 1
