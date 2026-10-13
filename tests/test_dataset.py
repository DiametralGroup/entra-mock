"""The dataset -- and above all its two edge cases.

This mock isn't worth its volume but what it lets us PROVE at the join with
the HR system. The four groups match the four rules of insights360's
authorization model, and two UPNs exist solely to exercise the edges:

  • `ext.consultant@boreal-conseil.example` -- member of a group, ABSENT from the HR system;
  • `kevin.silva@boreal-conseil.example`     -- present in HR, member of NO group.

If either one disappeared, the downstream authorization suite would pass
while no longer proving anything.
"""

from __future__ import annotations

from conftest import all_groups


def _all_members(client, auth) -> set[str]:
    members: set[str] = set()
    for group in all_groups(client, auth):
        url = f"/v1.0/groups/{group['id']}/members"
        while url:
            body = client.get(url, headers=auth).json()
            # `.get` and not `[...]`: not all members are people --
            # `members` returns `directoryObject`s, among which a nested
            # group and a service principal, neither with a UPN.
            members.update(
                m["userPrincipalName"] for m in body["value"] if m.get("userPrincipalName")
            )
            next_link = body.get("@odata.nextLink")
            url = next_link.replace("http://testserver", "") if next_link else ""
    return members


def test_the_four_groups_match_the_four_rules(client, auth):
    names = {g["displayName"] for g in all_groups(client, auth)}
    assert {"grp-bi-rh", "grp-bi-sales", "grp-bi-direction", "grp-bi-comex"} <= names


def test_the_nested_group_adds_no_membership(client, auth):
    """It exercises the consumer's DIAGNOSTIC, not its authorization model.

    `grp-bi-imbrique` contains only a group and a service principal: no UPN,
    hence no membership row. The downstream expectations are unchanged --
    what changes is that a consumer must now distinguish "object with no UPN
    expected" from "nested group, hence invisible rightful holders".
    """
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    members = client.get(f"/v1.0/groups/{groups['grp-bi-imbrique']}/members", headers=auth).json()
    assert all("userPrincipalName" not in m for m in members["value"])


def test_the_domain_matches_the_realiste_dataset(client, auth):
    """The UPN is the ONLY join key between the directory and the HR system.

    A domain that drifts from boondmanager-mock produces no error: it
    produces zero memberships, hence zero visibility, hence an authorization
    suite green BY VACUITY. This test is the guard rail -- it must break
    loudly the day one of the two datasets moves without the other.
    """
    members = _all_members(client, auth)
    assert members, "no member served"
    assert all(upn.endswith("@boreal-conseil.example") for upn in members), members


def test_the_member_absent_from_hr_is_present(client, auth):
    """It MUST be served: the pipeline must drop it at the join, never emit
    a row with an unknown key -- which would degrade the inner layer's
    predicate into "no filter"."""
    assert "ext.consultant@boreal-conseil.example" in _all_members(client, auth)


def test_the_collaborator_without_a_group_is_absent(client, auth):
    """`kevin.silva` belongs to NO group: he must see only himself (rule
    1). His presence here would invalidate the edge case."""
    assert "kevin.silva@boreal-conseil.example" not in _all_members(client, auth)


def test_comex_overlaps_direction(client, auth):
    """Rule 4 (full visibility) applies to a person who ALSO has a group
    scope: that overlap is what exercises the deduplication of the union of
    rules."""
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    comex = client.get(f"/v1.0/groups/{groups['grp-bi-comex']}/members", headers=auth).json()
    direction = client.get(
        f"/v1.0/groups/{groups['grp-bi-direction']}/members", headers=auth
    ).json()
    assert comex["value"][0]["userPrincipalName"] == direction["value"][0]["userPrincipalName"]


def test_the_dataset_is_deterministic(client, auth):
    """No randomness: two reads return the same world. The dataset is a
    FIXTURE, and downstream expectations rely on it."""
    assert _all_members(client, auth) == _all_members(client, auth)
