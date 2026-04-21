"""规则层测试。"""

from sap.rules import (
    build_fremdl_entry,
    build_lab_cost_entries,
    build_single_plan_cost_entries,
    build_split_plan_cost_entries,
    resolve_a2_materials,
    resolve_data_a_key,
)
from sap.test.helpers import make_config, make_cost_options, make_order, make_revenue


class TestResolveA2Materials:
    def test_use_mapping_by_subcode(self):
        config = make_config()
        assert resolve_a2_materials("A2-430-00", config) == ("T20-430-00", "T75-430-00")

    def test_use_default_mapping_when_unknown(self):
        config = make_config()
        assert resolve_a2_materials("A2-999-00", config) == ("T75-441-00", "T20-441-00")


class TestResolveDataAKey:
    def test_d2_in_data_ae1_returns_e1(self):
        config = make_config(data_ae1=["100001"])
        order = make_order(material_code="D2-405-00", sap_no="100001")
        assert resolve_data_a_key(order, config) == "E1"

    def test_d3_not_in_data_ae1_returns_z0(self):
        config = make_config(data_ae1=[])
        order = make_order(material_code="D3-405-00", sap_no="100001")
        assert resolve_data_a_key(order, config) == "Z0"

    def test_data_az2_returns_z2(self):
        config = make_config(data_az2=["200001"])
        order = make_order(material_code="T75-405-00", sap_no="200001")
        assert resolve_data_a_key(order, config) == "Z2"

    def test_default_returns_00(self):
        assert resolve_data_a_key(make_order(sap_no="999999"), make_config()) == "00"


class TestBuildLabCostEntries:
    def test_split_material_builds_two_entries(self):
        entries = build_lab_cost_entries(make_order(material_code="A2-405-00"), make_revenue(), make_config())
        assert len(entries) == 2
        assert entries[0].row == 0
        assert entries[1].row == 1

    def test_t20_builds_phy_entry(self):
        entries = build_lab_cost_entries(make_order(material_code="T20-441-00"), make_revenue(), make_config())
        assert len(entries) == 1
        assert entries[0].performer_cost_center == "1103"

    def test_t75_builds_chm_entry(self):
        entries = build_lab_cost_entries(make_order(material_code="T75-405-00"), make_revenue(), make_config())
        assert len(entries) == 1
        assert entries[0].performer_cost_center == "1102"


class TestPlanCostRules:
    def test_split_plan_cost_entries(self):
        entries = build_split_plan_cost_entries(make_revenue(), make_config(), make_cost_options())
        assert [entry.cost_center for entry in entries] == ["1101", "1102", "1103"]

    def test_single_plan_cost_entries_for_t75(self):
        entries = build_single_plan_cost_entries(
            make_order(material_code="T75-405-00"),
            make_revenue(cs_cost=1500.0, lab_cost=2500.0),
            make_config(),
            make_cost_options(),
        )
        assert [entry.cost_center for entry in entries] == ["1101", "1102"]

    def test_fremdl_entry_none_when_cost_zero(self):
        assert build_fremdl_entry(2, make_order(cost=0.0), make_config()) is None

    def test_fremdl_entry_created_when_cost_positive(self):
        entry = build_fremdl_entry(2, make_order(cost=500.0), make_config())
        assert entry is not None
        assert entry.row == 2
        assert entry.category == "FREMDL"
