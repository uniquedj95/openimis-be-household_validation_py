import math
from dataclasses import dataclass, field
from datetime import date
from functools import cached_property
from typing import Any


WEALTH_RANK = {
    "poorest": 1,
    "poorer": 2,
    "middle": 3,
    "richer": 4,
    "richest": 5,
}

CATEGORY_FEMALE_HEADED = "FEMALE_HEADED"
CATEGORY_YOUTH = "YOUTH"
# Every eligible household that is neither female-headed nor youth-headed
CATEGORY_OTHER = "OTHER"

ROW_TYPE_MAIN = "MAIN"
ROW_TYPE_RESERVE = "RESERVE"


@dataclass(frozen=True)
class EligibleMember:
    id: Any
    gender: str | None = None
    dob: date | None = None
    fit_for_work: bool = False
    role: str | None = None
    recipient_type: str | None = None
    data_source: str | None = None
    json_ext: dict = field(default_factory=dict)
    source: Any = None

    @property
    def age(self) -> int | None:
        if not self.dob:
            return None
        today = date.today()
        return today.year - self.dob.year - (
            (today.month, today.day) < (self.dob.month, self.dob.day)
        )

    def is_eligible(self, rule) -> bool:
        """Whether this member counts towards a household's eligible members.

        With no program-specific ``rule``, a member is eligible iff they're fit for work.
        A rule can instead require a specific ``data_source``, ``recipient_type``, a truthy
        ``json_ext`` flag, and/or an inclusive age range.
        """
        if not rule:
            return self.fit_for_work

        required_data_source = rule.get("requires_data_source")
        if required_data_source:
            if self.data_source is None:
                return False
            if str(self.data_source).strip().upper() != str(required_data_source).strip().upper():
                return False

        required_recipient_type = rule.get("requires_recipient_type")
        if required_recipient_type:
            if str(self.recipient_type or "").strip().upper() != str(required_recipient_type).strip().upper():
                return False

        flag = rule.get("member_flag")
        if flag and not is_truthy(self.json_ext.get(flag)):
            return False

        min_age = rule.get("member_min_age")
        if min_age is not None and (self.age is None or self.age < min_age):
            return False

        max_age = rule.get("member_max_age")
        if max_age is not None and (self.age is None or self.age > max_age):
            return False

        return True


@dataclass
class EligibleHousehold:
    id: Any
    code: str | None = None
    wealth_quintile: str | int | None = None
    pmt_score: str | float | None = None
    last_verified_date: date | None = None
    head: EligibleMember | None = None
    eligible_members: list[EligibleMember] = field(default_factory=list)
    source: Any = None
    village_id: Any = None
    village_code: str | None = None
    village_name: str | None = None


@dataclass(frozen=True)
class SelectedHousehold:
    household: EligibleHousehold
    category: str
    row_type: str


@dataclass(frozen=True)
class SelectedMember:
    household: EligibleHousehold
    member: EligibleMember
    category: str
    row_type: str


@dataclass(frozen=True)
class SelectionResult:
    main: list[SelectedHousehold]
    reserve: list[SelectedHousehold]

    @property
    def selected(self) -> list[SelectedHousehold]:
        return [*self.main, *self.reserve]

    @cached_property
    def member_rows(self) -> list[SelectedMember]:
        return [
            SelectedMember(
                household=selected.household,
                member=member,
                category=selected.category,
                row_type=selected.row_type,
            )
            for selected in self.selected
            for member in selected.household.eligible_members
        ]


def is_truthy(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y"}
    return bool(value)


def normalize_gender(value):
    if value is None:
        return None
    gender = str(value).strip().lower()
    if gender in {"f", "female", "woman"}:
        return "female"
    if gender in {"m", "male", "man"}:
        return "male"
    return gender


def normalize_wealth_rank(value):
    if value is None:
        return 999
    if isinstance(value, int):
        return value if value > 0 else 999
    if isinstance(value, float):
        return int(value) if value > 0 else 999
    text = str(value).strip().lower()
    if text.isdigit():
        number = int(text)
        return number if number > 0 else 999
    return WEALTH_RANK.get(text, 999)


def categorize_household(household: EligibleHousehold):
    if household.head and normalize_gender(household.head.gender) == "female":
        return CATEGORY_FEMALE_HEADED

    for member in household.eligible_members:
        if member.age is not None and 18 <= member.age <= 35:
            return CATEGORY_YOUTH

    return CATEGORY_OTHER


def household_sort_key(household: EligibleHousehold):
    return (
        normalize_wealth_rank(household.wealth_quintile),
        str(household.code or ""),
        str(household.id),
    )


def exclude_recently_verified(households, exclude_verified_after=None):
    if exclude_verified_after is None:
        return list(households)
    return [
        household
        for household in households
        if not household.last_verified_date
        or household.last_verified_date <= exclude_verified_after
    ]


def _quota_percentage(quota_config, key, default):
    value = (quota_config or {}).get(key)
    if value is None:
        return default
    try:
        value = int(value)
    except (TypeError, ValueError):
        return default
    return max(0, min(value, 100))


def female_headed_percentage(quota_config):
    return _quota_percentage(quota_config, "female_headed_percentage", 40)


def youth_headed_percentage(quota_config):
    return _quota_percentage(quota_config, "youth_headed_percentage", 40)


def reserve_percentage(quota_config):
    return _quota_percentage(quota_config, "reserve_percentage", 20)


def _allocate_quotas(target_count, quota_config):
    female_weight = female_headed_percentage(quota_config) / 100
    youth_weight = youth_headed_percentage(quota_config) / 100
    if female_weight + youth_weight > 1:
        total = female_weight + youth_weight
        female_weight = female_weight / total
        youth_weight = youth_weight / total
    other_weight = max(0, 1 - female_weight - youth_weight)
    weights = [
        (CATEGORY_FEMALE_HEADED, female_weight),
        (CATEGORY_YOUTH, youth_weight),
        (CATEGORY_OTHER, other_weight),
    ]
    exact = [(category, target_count * weight) for category, weight in weights]
    quotas = {category: math.floor(value) for category, value in exact}
    remaining = target_count - sum(quotas.values())
    remainders = sorted(
        exact,
        key=lambda item: (item[1] - math.floor(item[1])),
        reverse=True,
    )
    for category, _value in remainders[:remaining]:
        quotas[category] += 1
    return quotas


def _village_key(household):
    if household.village_code:
        return f"code:{household.village_code}"
    if household.village_id is not None:
        return f"id:{household.village_id}"
    return "unassigned"


def _allocate_village_targets(households_by_village, target_count):
    """Allocate an exact target proportionally, with one slot per village when possible."""
    capacities = {
        key: len(households)
        for key, households in households_by_village.items()
        if households
    }
    target_count = max(0, min(target_count, sum(capacities.values())))
    if not capacities or target_count == 0:
        return {key: 0 for key in capacities}, {key: 0.0 for key in capacities}

    total_capacity = sum(capacities.values())
    exact = {
        key: target_count * capacity / total_capacity
        for key, capacity in capacities.items()
    }
    allocations = {
        key: min(capacity, math.floor(exact[key]))
        for key, capacity in capacities.items()
    }
    remaining = target_count - sum(allocations.values())
    remainder_order = sorted(
        capacities,
        key=lambda key: (-(exact[key] - math.floor(exact[key])), str(key)),
    )
    while remaining:
        allocated_in_pass = False
        for key in remainder_order:
            if allocations[key] >= capacities[key]:
                continue
            allocations[key] += 1
            remaining -= 1
            allocated_in_pass = True
            if remaining == 0:
                break
        if not allocated_in_pass:
            break

    # When there are enough target slots, guarantee representation for every
    # village containing an eligible household. Preserve the exact total by
    # moving a slot from the strongest donor rather than adding a new slot.
    if target_count >= len(capacities):
        for empty_key in sorted(
            (key for key, value in allocations.items() if value == 0),
            key=str,
        ):
            donors = [key for key, value in allocations.items() if value > 1]
            if not donors:
                break
            donor_key = max(
                donors,
                key=lambda key: (
                    allocations[key] - exact[key],
                    allocations[key],
                    str(key),
                ),
            )
            allocations[donor_key] -= 1
            allocations[empty_key] += 1

    return allocations, exact


def _select_main_households(eligible, main_target, quota_config):
    quotas = _allocate_quotas(main_target, quota_config)
    by_category = {
        CATEGORY_FEMALE_HEADED: [],
        CATEGORY_YOUTH: [],
        CATEGORY_OTHER: [],
    }
    household_categories = {}
    for household in eligible:
        category = categorize_household(household)
        household_categories[household.id] = category
        by_category[category].append(household)

    selected_ids = set()
    main = []
    category_counts = {
        CATEGORY_FEMALE_HEADED: 0,
        CATEGORY_YOUTH: 0,
        CATEGORY_OTHER: 0,
    }
    selected_individuals = 0
    for category in (CATEGORY_FEMALE_HEADED, CATEGORY_YOUTH, CATEGORY_OTHER):
        for household in by_category[category]:
            if len(main) >= main_target:
                break
            if category_counts[category] >= quotas[category]:
                break
            selected_ids.add(household.id)
            category_counts[category] += 1
            selected_individuals += len(household.eligible_members)
            main.append(SelectedHousehold(household, category, ROW_TYPE_MAIN))

    if len(main) < main_target:
        for household in eligible:
            if household.id in selected_ids:
                continue
            selected_ids.add(household.id)
            category = household_categories[household.id]
            category_counts[category] += 1
            selected_individuals += len(household.eligible_members)
            main.append(SelectedHousehold(household, category, ROW_TYPE_MAIN))
            if len(main) >= main_target:
                break

    return (
        main,
        selected_ids,
        category_counts,
        selected_individuals,
        household_categories,
    )


def _eligible_pool(households, exclude_verified_after):
    """Households eligible for either selection algorithm: not recently
    re-verified, and with at least one eligible member."""
    return [
        household
        for household in exclude_recently_verified(households, exclude_verified_after)
        if household.eligible_members
    ]


def _cap_target(target_count, pool_size):
    """Clamp a requested ``target_count`` into ``[0, pool_size]``; ``None`` means "all"."""
    target = pool_size if target_count is None else target_count
    return max(0, min(target, pool_size))


def _build_summary(main, reserve, selected_individuals, category_counts, village_breakdown):
    return {
        "selected_households": len(main),
        "selected_individuals": selected_individuals,
        "selected_female_headed_households": category_counts[CATEGORY_FEMALE_HEADED],
        "selected_youth_households": category_counts[CATEGORY_YOUTH],
        "selected_other_households": category_counts[CATEGORY_OTHER],
        "reserve_households": len(reserve),
        "village_breakdown": village_breakdown,
    }


def select_households(
    households,
    target_count=None,
    exclude_verified_after=None,
    allocate_by_village=False,
    rule=None,
):
    """Single selection entry point for both the default PWP algorithm and
    program-based (e.g. Jobs Now) instances.

    ``rule`` is resolved from ``program_eligibility_rules`` for the selected
    Program (falling back to its ``"PWP"`` entry when no Program is selected
    or none matches). Whether its ``selection_strategy`` key is present picks
    the algorithm:

    - Present, as a dict (PWP's default): wealth-ranked, female-headed/
      youth/other demographic-quota allocation — percentages read from that
      dict's ``female_headed_percentage``/``youth_headed_percentage``/
      ``reserve_percentage`` — optionally allocated proportionally by
      village, with a reserve list.
    - Absent (e.g. Jobs Now's RMEP/UPG): no wealth ranking and no
      demographic quotas — every eligible household is selected, ordered by
      the rule's ``priority_flag`` (if any) and capped at ``target_count``
      with no reserve list. ``allocate_by_village`` has no effect in this
      mode: program-based generation has no hotspot/micro-catchment concept
      to scope a village allocation by.
    """
    quota_config = (rule or {}).get("selection_strategy")

    eligible = _eligible_pool(households, exclude_verified_after)
    eligible = sorted(
        eligible,
        key=household_sort_key if quota_config is not None else _eligible_household_sort_key(rule),
    )

    main_target = _cap_target(target_count, len(eligible))

    if quota_config is None:
        selected = eligible[:main_target]
        main = [
            SelectedHousehold(household, CATEGORY_OTHER, ROW_TYPE_MAIN)
            for household in selected
        ]
        reserve = []
        selected_individuals = sum(len(household.eligible_members) for household in selected)
        category_counts = {
            CATEGORY_FEMALE_HEADED: 0,
            CATEGORY_YOUTH: 0,
            CATEGORY_OTHER: len(main),
        }
        village_breakdown = []
        selection_result = SelectionResult(main=main, reserve=reserve)
        summary = _build_summary(main, reserve, selected_individuals, category_counts, village_breakdown)
        return selection_result, summary

    requested_reserve_target = math.ceil(
        main_target * reserve_percentage(quota_config) / 100
    )

    village_breakdown = []
    if allocate_by_village and target_count is not None:
        households_by_village = {}
        for household in eligible:
            households_by_village.setdefault(_village_key(household), []).append(household)

        main_allocations, exact_allocations = _allocate_village_targets(
            households_by_village,
            main_target,
        )
        main = []
        selected_ids = set()
        selected_individuals = 0
        category_counts = {
            CATEGORY_FEMALE_HEADED: 0,
            CATEGORY_YOUTH: 0,
            CATEGORY_OTHER: 0,
        }
        household_categories = {
            household.id: categorize_household(household)
            for household in eligible
        }
        village_selected_individuals = {}
        for key in sorted(households_by_village, key=str):
            village_main, village_ids, village_counts, village_individuals, _ = (
                _select_main_households(
                    households_by_village[key],
                    main_allocations.get(key, 0),
                    quota_config,
                )
            )
            main.extend(village_main)
            selected_ids.update(village_ids)
            selected_individuals += village_individuals
            village_selected_individuals[key] = village_individuals
            for category, count in village_counts.items():
                category_counts[category] += count

        reserve_target = max(
            0,
            min(
                requested_reserve_target,
                len(eligible) - len(selected_ids),
            ),
        )
        remaining_by_village = {
            key: [
                household
                for household in households
                if household.id not in selected_ids
            ]
            for key, households in households_by_village.items()
        }
        reserve_allocations, _reserve_exact = _allocate_village_targets(
            remaining_by_village,
            reserve_target,
        )
        reserve = []
        for key in sorted(remaining_by_village, key=str):
            for household in remaining_by_village[key][:reserve_allocations.get(key, 0)]:
                selected_ids.add(household.id)
                reserve.append(
                    SelectedHousehold(
                        household,
                        household_categories[household.id],
                        ROW_TYPE_RESERVE,
                    )
                )

        for key in sorted(households_by_village, key=str):
            representative = households_by_village[key][0]
            allocation = main_allocations.get(key, 0)
            village_breakdown.append(
                {
                    "village_id": representative.village_id,
                    "village_code": representative.village_code,
                    "village_name": representative.village_name,
                    "eligible_households": len(households_by_village[key]),
                    "exact_allocation": round(exact_allocations.get(key, 0), 6),
                    "allocated_households": allocation,
                    "selected_households": allocation,
                    "selected_individuals": village_selected_individuals.get(key, 0),
                    "reserve_households": reserve_allocations.get(key, 0),
                }
            )
    else:
        (
            main,
            selected_ids,
            category_counts,
            selected_individuals,
            household_categories,
        ) = _select_main_households(eligible, main_target, quota_config)
        reserve_target = max(
            0,
            min(
                requested_reserve_target,
                len(eligible) - len(selected_ids),
            ),
        )
        reserve = []
        for household in eligible:
            if household.id in selected_ids:
                continue
            selected_ids.add(household.id)
            reserve.append(
                SelectedHousehold(
                    household,
                    household_categories[household.id],
                    ROW_TYPE_RESERVE,
                )
            )
            if len(reserve) >= reserve_target:
                break

    selection_result = SelectionResult(main=main, reserve=reserve)
    summary = _build_summary(main, reserve, selected_individuals, category_counts, village_breakdown)
    return selection_result, summary


def _household_has_priority_flag(household: EligibleHousehold, priority_flag):
    if not priority_flag:
        return False
    return any(
        is_truthy(member.json_ext.get(priority_flag))
        for member in household.eligible_members
    )


def _eligible_household_sort_key(rule):
    """Sort key factory for the ``rule``-driven branch of ``select_households``.

    Households with an eligible member carrying the rule's ``priority_flag``
    sort first; everything else keeps a stable, deterministic order behind them.
    """
    priority_flag = (rule or {}).get("priority_flag")

    def key(household):
        has_priority = _household_has_priority_flag(household, priority_flag)
        return (not has_priority, str(household.code or ""), str(household.id))

    return key
