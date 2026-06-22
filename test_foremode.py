"""Self-checks for the pFMEA engine. Run: python test_pfmea.py  (or pytest)."""
import foremode


def test_action_priority_known_cases():
    assert foremode.get_action_priority(8, 4, 3)[0] == "L"
    assert foremode.get_action_priority(9, 3, 2)[0] == "M"
    assert foremode.get_action_priority(7, 4, 7)[0] == "H"
    assert foremode.get_action_priority(10, 2, 8)[0] == "H"   # S>=9 + D>=7
    assert foremode.get_action_priority(3, 2, 2)[0] == "L"


def test_action_priority_validates_range():
    for bad in [(11, 1, 1), (5, 0, 5), (5, 5, 11)]:
        try:
            foremode.get_action_priority(*bad)
            assert False, f"expected ValueError for {bad}"
        except ValueError:
            pass


def test_metrics_consistent_with_ap():
    sc, _ = foremode.load_scenario("cnc_femoral_stem")
    metrics = dict(foremode.compute_metrics(sc))
    initial_h = sum(1 for r in sc["ratings"]
                    if foremode.get_action_priority(r[1], r[2], r[3])[0] == "H")
    assert metrics["Initial High (H) Action Priority"] == initial_h
    assert metrics["Total Failure Modes Analyzed"] == len(sc["ratings"])


def test_all_scenarios_build():
    for sid in ("cnc_femoral_stem", "spinal_peek_cage", "sterile_packaging"):
        sc, _ = foremode.load_scenario(sid)
        wb = foremode.build_workbook(sc, iso14971=True)
        assert len(wb.worksheets) == 8   # 7 steps + ISO 14971 bridge


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); print(f"ok  {name}")
    print("all checks passed")
