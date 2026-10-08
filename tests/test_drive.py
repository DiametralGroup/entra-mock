"""La surface « fichiers » de Graph (driveItem).

Chaque test vise un comportement du VRAI Graph qu'un mock gentil cacherait :
le 403 d'une application `Sites.Selected` sans droit sur le site, la
pagination de `/children`, le jeu de propriétés par défaut qui porte une
identité, l'empreinte `quickXorHash` seule, la redirection 302 vers une URL
pré-authentifiée à suivre SANS le Bearer.

Le lecteur est VIDE au démarrage : les tests qui ont besoin de fichiers les
écrivent eux-mêmes par le plan de contrôle — de petits fichiers neutres.
"""

from __future__ import annotations

import base64
import re
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlsplit

import pytest
import yaml
from conftest import ADMIN, app_module, drive_module

SITE = "/v1.0/sites/contoso.sharepoint.com:/sites/documents"
DRIVE = drive_module.DRIVE_ID
RACINE = f"/v1.0/drives/{DRIVE}/root/children"
REPORTS = f"/v1.0/drives/{DRIVE}/root:/reports:/children"

#: Un petit arbre neutre : six enfants sous `reports/` (trois pages de deux),
#: un sous-dossier, un fichier à la racine.
ARBRE = {
    "reports/report-1.txt": b"hello",
    "reports/report-2.txt": b"world",
    "reports/report-3.csv": b"a,b\n1,2\n",
    "reports/report-4.txt": b"four",
    "reports/report-5.txt": b"five",
    "reports/archive/old.txt": b"old",
    "notes.txt": b"notes",
}


@pytest.fixture(autouse=True)
def lecteur_vide():
    """Un lecteur REMIS À VIDE — avant et après, pour qu'une écriture ou un
    droit retiré ne fuie pas d'un test à l'autre."""
    drive_module.lecteur_mock.reinitialiser()
    yield
    drive_module.lecteur_mock.reinitialiser()


def ecrire(client, chemin: str, contenu: bytes):
    return client.put(f"/__admin/drive/files/{chemin}", content=contenu, headers=ADMIN)


@pytest.fixture()
def arbre(client) -> None:
    for chemin, contenu in ARBRE.items():
        assert ecrire(client, chemin, contenu).status_code == 201, chemin


def _relatif(url: str) -> str:
    return url.replace("http://testserver", "")


def tous_les_enfants(client, auth, url: str) -> list[dict[str, Any]]:
    """Les enfants en SUIVANT `@odata.nextLink` jusqu'au bout."""
    elements: list[dict[str, Any]] = []
    while url:
        reponse = client.get(_relatif(url), headers=auth)
        assert reponse.status_code == 200, reponse.text
        corps = reponse.json()
        elements.extend(corps["value"])
        url = corps.get("@odata.nextLink", "")
    return elements


def tout_l_arbre(client, auth, select: str = "") -> dict[str, dict[str, Any]]:
    """L'arborescence entière, chemin → élément : descendre dans CHAQUE dossier
    (par identifiant), en suivant la pagination à CHAQUE niveau."""
    suffixe = f"?$select={select}" if select else ""
    rendu: dict[str, dict[str, Any]] = {}
    a_visiter = [("", RACINE)]
    while a_visiter:
        prefixe, url = a_visiter.pop()
        for item in tous_les_enfants(client, auth, url + suffixe):
            chemin = prefixe + item["name"]
            rendu[chemin] = item
            if "folder" in item:
                enfants = f"/v1.0/drives/{DRIVE}/items/{item['id']}/children"
                a_visiter.append((f"{chemin}/", enfants))
    return rendu


def element(client, auth, chemin: str) -> dict[str, Any]:
    reponse = client.get(f"/v1.0/drives/{DRIVE}/root:/{chemin}", headers=auth)
    assert reponse.status_code == 200, reponse.text
    return dict(reponse.json())


def pages(client, auth, url: str) -> list[list[str]]:
    """Les noms, PAGE par page, en suivant le curseur."""
    rendu: list[list[str]] = []
    while url:
        corps = client.get(_relatif(url), headers=auth).json()
        rendu.append([e["name"] for e in corps["value"]])
        url = corps.get("@odata.nextLink", "")
    return rendu


def telecharger(client, auth, item_id: str) -> bytes:
    """Le geste CORRECT : `/content` sans suivre, puis `Location` SANS jeton."""
    redirection = client.get(
        f"/v1.0/drives/{DRIVE}/items/{item_id}/content", headers=auth, follow_redirects=False
    )
    assert redirection.status_code == 302, redirection.text
    reponse = client.get(redirection.headers["location"])
    assert reponse.status_code == 200, reponse.text
    return bytes(reponse.content)


# ── Le site et ses bibliothèques ─────────────────────────────────────────────


def test_le_site_par_son_chemin_rend_un_identifiant_composite(client, auth):
    r = client.get(SITE, headers=auth)
    assert r.status_code == 200, r.text
    site = r.json()
    hote, collection, web = site["id"].split(",")
    assert hote == "contoso.sharepoint.com"
    assert re.fullmatch(r"[0-9a-f-]{36}", collection) and re.fullmatch(r"[0-9a-f-]{36}", web)
    assert site["webUrl"] == "https://contoso.sharepoint.com/sites/documents"
    assert {"displayName", "name"} <= set(site)
    # Le même site par son identifiant.
    assert client.get(f"/v1.0/sites/{site['id']}", headers=auth).json()["id"] == site["id"]


@pytest.mark.parametrize(
    "chemin",
    [
        "/v1.0/sites/contoso.sharepoint.com:/sites/documents:",
        "/v1.0/sites/contoso.sharepoint.com:/sites/Documents",
        "/v1.0/sites/contoso.sharepoint.com%3A/sites/documents",
    ],
)
def test_les_variantes_d_adressage_du_site(client, auth, chemin):
    """Deux-points final, casse, deux-points encodé : Graph accepte les trois."""
    assert client.get(chemin, headers=auth).status_code == 200


def test_site_inconnu_404_itemNotFound(client, auth):
    r = client.get("/v1.0/sites/contoso.sharepoint.com:/sites/autre", headers=auth)
    assert r.status_code == 404
    assert r.json()["error"]["code"] == "itemNotFound"
    assert client.get("/v1.0/sites/inconnu,a,b", headers=auth).status_code == 404


def test_hote_d_un_autre_locataire_400(client, auth):
    r = client.get("/v1.0/sites/fabrikam.sharepoint.com:/sites/documents", headers=auth)
    assert r.status_code == 400
    assert r.json()["error"]["code"] == "invalidRequest"


def test_sans_droit_sur_le_site_403_accessDenied_PARTOUT(client, auth, monkeypatch):
    """L'application `Sites.Selected` à qui personne n'a accordé le site.

    Le jeton est VALIDE — c'est le site qui est fermé. Le 403 doit tomber sur
    le site comme sur tout ce qu'il contient, avec l'enveloppe de Graph : c'est
    ce qui permet au client de dire « demander le droit » plutôt que
    « renouveler le jeton » (401) ou « corriger le chemin » (404).
    """
    monkeypatch.setattr(drive_module, "SITE_GRANT", "none")
    drive_module.lecteur_mock.reinitialiser()
    ecrire(client, "reports/report-1.txt", b"hello")
    for url in (
        SITE,
        f"/v1.0/sites/{drive_module.SITE_ID}/drive",
        RACINE,
        REPORTS,
        f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
    ):
        r = client.get(url, headers=auth, follow_redirects=False)
        assert r.status_code == 403, url
        erreur = r.json()["error"]
        assert erreur["code"] == "accessDenied"
        assert erreur["message"] == "Access denied"
        assert {"date", "request-id", "client-request-id"} <= set(erreur["innerError"])


def test_le_droit_se_retire_et_se_rend_a_chaud(client, auth):
    """Un conteneur de CI ne se redémarre pas entre deux tests : le pendant
    runtime de ENTRA_MOCK_SITE_GRANT passe par le plan de contrôle."""
    assert client.post("/__admin/drive/grant", json={"grant": "none"}, headers=ADMIN).is_success
    assert client.get(SITE, headers=auth).status_code == 403
    assert client.post("/__admin/drive/grant", json={"grant": "read"}, headers=ADMIN).is_success
    assert client.get(SITE, headers=auth).status_code == 200
    r = client.post("/__admin/drive/grant", json={"grant": "owner"}, headers=ADMIN)
    assert r.status_code == 422


def test_la_bibliotheque_par_defaut_et_la_liste(client, auth):
    site_id = client.get(SITE, headers=auth).json()["id"]
    lecteur = client.get(f"/v1.0/sites/{site_id}/drive", headers=auth).json()
    assert lecteur["name"] == "Documents"
    assert lecteur["driveType"] == "documentLibrary"
    assert lecteur["id"].startswith("b!") and len(lecteur["id"]) == 66
    lecteurs = client.get(f"/v1.0/sites/{site_id}/drives", headers=auth).json()
    assert [d["id"] for d in lecteurs["value"]] == [lecteur["id"]]
    assert client.get(f"/v1.0/drives/{lecteur['id']}", headers=auth).json()["name"] == "Documents"


# ── Le lecteur VIDE par défaut ───────────────────────────────────────────────


def test_le_lecteur_est_VIDE_au_demarrage(client, auth):
    """Aucun fichier, aucun dossier : la racine seule."""
    corps = client.get(RACINE, headers=auth).json()
    assert corps["value"] == []
    assert "@odata.nextLink" not in corps
    racine = client.get(f"/v1.0/drives/{DRIVE}/root", headers=auth).json()
    assert racine["folder"] == {"childCount": 0}
    assert racine["root"] == {}
    assert racine["size"] == 0
    assert racine["name"] == "root"


# ── Les enfants d'un dossier ─────────────────────────────────────────────────


@pytest.mark.usefixtures("arbre")
def test_un_dossier_PAGINE_par_curseur_opaque(client, auth):
    """Six enfants, deux par page : un client qui ignore `@odata.nextLink` ne
    voit que les deux premiers — sans la moindre erreur."""
    premiere = client.get(REPORTS, headers=auth).json()
    assert len(premiere["value"]) == 2
    suivant = premiere["@odata.nextLink"]
    assert suivant.startswith("http://testserver/v1.0/drives/"), "nextLink doit être ABSOLU"
    jeton = parse_qs(urlsplit(suivant).query)["$skiptoken"][0]
    assert not jeton.isdigit(), "le curseur doit être OPAQUE, pas un rang"

    assert pages(client, auth, REPORTS) == [
        ["archive", "report-1.txt"],
        ["report-2.txt", "report-3.csv"],
        ["report-4.txt", "report-5.txt"],
    ]


@pytest.mark.usefixtures("arbre")
def test_par_chemin_et_par_identifiant_la_meme_pagination(client, auth):
    """`root:/reports:/children` et `items/{id}/children` : mêmes pages, et
    chaque lien suivant garde la forme de la requête."""
    reports = element(client, auth, "reports")
    par_id = f"/v1.0/drives/{DRIVE}/items/{reports['id']}/children"
    assert pages(client, auth, REPORTS) == pages(client, auth, par_id)
    assert (
        "/root:/reports:/children?" in client.get(REPORTS, headers=auth).json()["@odata.nextLink"]
    )
    assert (
        f"/items/{reports['id']}/children?"
        in client.get(par_id, headers=auth).json()["@odata.nextLink"]
    )


@pytest.mark.usefixtures("arbre")
def test_tout_l_arbre(client, auth):
    rendu = tout_l_arbre(client, auth)
    assert {c for c, e in rendu.items() if "file" in e} == set(ARBRE)
    assert {c for c, e in rendu.items() if "folder" in e} == {"reports", "reports/archive"}
    assert rendu["reports"]["folder"]["childCount"] == 6
    assert rendu["reports"]["size"] == sum(len(v) for k, v in ARBRE.items() if k.startswith("rep"))


@pytest.mark.usefixtures("arbre")
def test_top_est_un_plafond_pas_une_promesse(client, auth):
    """`$top=500` ne court-circuite PAS la pagination : le serveur rend moins."""
    grande = client.get(f"{REPORTS}?$top=500", headers=auth).json()
    assert len(grande["value"]) == 2
    assert "$top=500" in grande["@odata.nextLink"]
    assert len(client.get(f"{REPORTS}?$top=1", headers=auth).json()["value"]) == 1
    assert client.get(f"{REPORTS}?$top=zero", headers=auth).status_code == 400


def test_un_curseur_illisible_400(client, auth):
    assert client.get(f"{RACINE}?$skiptoken=pas-un-jeton", headers=auth).status_code == 400


@pytest.mark.usefixtures("arbre")
def test_le_select_MINIMISE_et_suit_le_lien(client, auth):
    """`$select` rend EXACTEMENT les propriétés demandées — pas d'`id` offert,
    pas d'identité — et il est recopié dans le lien suivant."""
    demande = "id,name,size,cTag,lastModifiedDateTime,file,folder"
    rendu = tout_l_arbre(client, auth, demande)
    assert len(rendu) == len(ARBRE) + 2
    for item in rendu.values():
        assert set(item) <= set(demande.split(",")), item
        assert "createdBy" not in item and "lastModifiedBy" not in item
    premiere = client.get(f"{REPORTS}?$select=name", headers=auth).json()
    assert "$select=name" in premiere["@odata.nextLink"]
    assert all(
        set(item) == {"name"} for item in tous_les_enfants(client, auth, f"{REPORTS}?$select=name")
    )


def test_un_select_inconnu_400(client, auth):
    r = client.get(f"{RACINE}?$select=name,sha256Hash", headers=auth)
    assert r.status_code == 400
    assert r.json()["error"]["code"] == "BadRequest"
    assert "sha256Hash" in r.json()["error"]["message"]
    assert client.get(f"{SITE}?$select=nope", headers=auth).status_code == 400


def test_sans_select_le_jeu_par_defaut_expose_une_identite(client, auth):
    """Le défaut que `$select` existe pour éviter : un nom et un courriel."""
    ecrire(client, "reports/report-1.txt", b"hello")
    item = element(client, auth, "reports/report-1.txt")
    for cle in ("createdBy", "lastModifiedBy"):
        assert item[cle]["user"]["displayName"] == "Mock User"
        assert item[cle]["user"]["email"] == "user@example.invalid"
    assert {"parentReference", "webUrl", "fileSystemInfo", "@microsoft.graph.downloadUrl"} <= set(
        item
    )
    assert item["parentReference"]["path"] == f"/drives/{DRIVE}/root:/reports"


@pytest.mark.usefixtures("arbre")
def test_SEUL_quickXorHash_jamais_sha(client, auth):
    for item in tout_l_arbre(client, auth).values():
        if "folder" in item:
            assert "file" not in item
            continue
        assert set(item["file"]["hashes"]) == {"quickXorHash"}
        assert "folder" not in item
    assert element(client, auth, "reports/report-3.csv")["file"]["mimeType"] == "text/csv"


def test_etag_et_ctag_ont_la_forme_attendue(client, auth):
    ecrire(client, "reports/report-1.txt", b"hello")
    item = element(client, auth, "reports/report-1.txt")
    assert re.fullmatch(r'"\{[0-9A-F-]{36}\},1"', item["eTag"])
    assert re.fullmatch(r'"c:\{[0-9A-F-]{36}\},1"', item["cTag"])
    assert item["eTag"][2:38] == item["cTag"][4:40]


@pytest.mark.usefixtures("arbre")
def test_element_par_identifiant_egal_element_par_chemin(client, auth):
    par_chemin = element(client, auth, "reports/report-2.txt")
    par_id = client.get(f"/v1.0/drives/{DRIVE}/items/{par_chemin['id']}", headers=auth).json()
    for cle in ("id", "name", "size", "eTag", "cTag", "file", "lastModifiedDateTime"):
        assert par_id[cle] == par_chemin[cle]
    assert re.fullmatch(r"01[A-Z2-7]{32}", par_id["id"])


@pytest.mark.usefixtures("arbre")
def test_la_casse_du_chemin_est_indifferente(client, auth):
    assert client.get(REPORTS.replace("reports", "REPORTS"), headers=auth).status_code == 200


def test_dossier_inconnu_404_itemNotFound(client, auth):
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/absent:/children", headers=auth)
    assert r.status_code == 404
    assert r.json()["error"]["code"] == "itemNotFound"
    assert client.get("/v1.0/drives/b!inconnu/root/children", headers=auth).status_code == 404
    assert client.get(f"/v1.0/drives/{DRIVE}/items/01INCONNU", headers=auth).status_code == 404
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/absent:/inconnu", headers=auth)
    assert r.status_code == 400


# ── Le contenu : 302, puis une URL pré-authentifiée ──────────────────────────


def test_content_rend_302_vers_une_url_pre_authentifiee(client, auth):
    ecrire(client, "reports/report-1.txt", b"hello")
    item = element(client, auth, "reports/report-1.txt")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/items/{item['id']}/content", headers=auth, follow_redirects=False
    )
    assert r.status_code == 302
    lieu = urlsplit(r.headers["location"])
    assert lieu.scheme == "http" and lieu.netloc == "testserver", "Location doit être ABSOLUE"
    assert lieu.path == "/sites/documents/_layouts/15/download.aspx"
    assert {"UniqueId", "tempauth"} <= set(parse_qs(lieu.query))

    # Suivie SANS jeton : les octets, avec le bon type.
    telechargement = client.get(r.headers["location"])
    assert telechargement.status_code == 200
    assert telechargement.content == b"hello"
    assert telechargement.headers["content-type"].startswith("text/plain")
    assert len(telechargement.content) == item["size"]


def test_content_par_chemin_aussi_et_pas_sur_un_dossier(client, auth):
    ecrire(client, "reports/report-1.txt", b"hello")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
        headers=auth,
        follow_redirects=False,
    )
    assert r.status_code == 302
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/reports:/content", headers=auth)
    assert r.status_code == 404


def test_renvoyer_le_Bearer_a_l_url_de_telechargement_401(client, auth):
    """LA faute : l'autorisation est DANS l'URL. Le jeton Graph n'a rien à y
    faire — un client qui le renvoie est refusé."""
    ecrire(client, "reports/report-1.txt", b"hello")
    item = element(client, auth, "reports/report-1.txt")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/items/{item['id']}/content", headers=auth, follow_redirects=False
    )
    assert client.get(r.headers["location"], headers=auth).status_code == 401
    # Le piège en vrai : suivre la redirection AVEC le jeton (même origine ici).
    suivie = client.get(f"/v1.0/drives/{DRIVE}/items/{item['id']}/content", headers=auth)
    assert suivie.status_code == 401


def test_tempauth_falsifie_inconnu_ou_expire_401(client, auth, monkeypatch):
    ecrire(client, "reports/report-1.txt", b"hello")
    ecrire(client, "reports/report-2.txt", b"world")
    url = element(client, auth, "reports/report-1.txt")["@microsoft.graph.downloadUrl"]
    assert client.get(url).status_code == 200
    jeton = parse_qs(urlsplit(url).query)["tempauth"][0]
    falsifie = url.replace(jeton, jeton[:-4] + ("AAAA" if not jeton.endswith("AAAA") else "BBBB"))
    assert client.get(falsifie).status_code == 401
    assert client.get(url.replace(jeton, "inconnu")).status_code == 401
    # Un tempauth valide pour un AUTRE fichier ne l'ouvre pas.
    autre = element(client, auth, "reports/report-2.txt")["@microsoft.graph.downloadUrl"]
    jeton_autre = parse_qs(urlsplit(autre).query)["tempauth"][0]
    assert client.get(url.replace(jeton, jeton_autre)).status_code == 401

    monkeypatch.setattr(drive_module, "TEMPAUTH_SECONDS", -1)
    expire = element(client, auth, "reports/report-1.txt")["@microsoft.graph.downloadUrl"]
    assert client.get(expire).status_code == 401


def test_quick_xor_hash_vecteurs_de_reference():
    """Vecteurs relevés contre l'implémentation de référence (paquet PyPI
    `quickxorhash` 1.0.5, le 2026-10-08)."""
    vecteurs = {
        b"": "AAAAAAAAAAAAAAAAAAAAAAAAAAA=",
        b"a": "YQAAAAAAAAAAAAAAAQAAAAAAAAA=",
        b"hello world": "aCgDG9jwBhDc4Q1yawMZAAAAAAA=",
        bytes(range(256)) * 3: "rxAOGe1RimTF/e+k/m0O5nnSZT8=",
    }
    for donnees, attendu in vecteurs.items():
        assert drive_module.quick_xor_hash(donnees) == attendu


def test_le_quickXorHash_servi_est_celui_des_octets(client, auth):
    contenu = bytes(range(256)) * 3
    ecrire(client, "data/blob.bin", contenu)
    item = element(client, auth, "data/blob.bin")
    assert telecharger(client, auth, item["id"]) == contenu
    assert item["file"]["hashes"]["quickXorHash"] == "rxAOGe1RimTF/e+k/m0O5nnSZT8="
    assert len(base64.b64decode(item["file"]["hashes"]["quickXorHash"])) == 20
    assert item["file"]["mimeType"] == "application/octet-stream"


# ── Le plan de contrôle ──────────────────────────────────────────────────────


def test_le_plan_de_controle_exige_son_jeton(client):
    assert client.post("/__admin/drive/reset").status_code == 403
    assert (
        client.post("/__admin/drive/reset", headers={"X-Mock-Admin-Token": "faux"}).status_code
        == 403
    )
    assert client.put("/__admin/drive/files/a.txt", content=b"x").status_code == 403
    assert client.get("/__admin/drive/counters").status_code == 403
    assert client.post("/__admin/drive/reset", headers=ADMIN).status_code == 200


def test_ecraser_GARDE_l_id_et_avance_ctag_etag_et_l_horloge(client, auth):
    """Un fichier réécrit sous le MÊME nom : même élément (même id), contenu
    neuf, `n` passe à 2 dans `eTag` et `cTag`, `lastModifiedDateTime` avance —
    sur l'horloge du mock, pas l'horloge murale."""
    assert ecrire(client, "reports/report-1.txt", b"hello").status_code == 201
    avant = element(client, auth, "reports/report-1.txt")
    assert avant["createdDateTime"] == "2026-01-01T00:02:00Z"  # le dossier a pris 00:01
    r = ecrire(client, "reports/report-1.txt", b"hello, again")
    assert r.status_code == 200, r.text
    apres = element(client, auth, "reports/report-1.txt")

    assert apres["id"] == avant["id"]
    assert apres["cTag"] == avant["cTag"].replace(",1", ",2")
    assert apres["eTag"] == avant["eTag"].replace(",1", ",2")
    assert apres["lastModifiedDateTime"] == "2026-01-01T00:03:00Z"
    assert apres["createdDateTime"] == avant["createdDateTime"]
    assert apres["file"]["hashes"] != avant["file"]["hashes"]
    assert apres["size"] == len(b"hello, again")
    assert telecharger(client, auth, apres["id"]) == b"hello, again"


def test_ecrire_cree_les_dossiers_intermediaires(client, auth):
    r = ecrire(client, "a/b/c/deep.txt", b"x")
    assert r.status_code == 201, r.text
    assert r.json()["parentReference"]["path"].endswith("/root:/a/b/c")
    assert r.json()["lastModifiedDateTime"] == "2026-01-01T00:04:00Z"
    for i, dossier in enumerate(("a", "a/b", "a/b/c"), start=1):
        item = element(client, auth, dossier)
        assert item["folder"]["childCount"] == 1
        assert item["createdDateTime"] == f"2026-01-01T00:0{i}:00Z"
    assert [e["name"] for e in tous_les_enfants(client, auth, RACINE)] == ["a"]


def test_le_mock_n_utilise_aucune_horloge_murale(client):
    """La même suite d'écritures rend les mêmes horodatages ET les mêmes
    identifiants, run après run."""

    def releve() -> list[tuple[str, str]]:
        return [
            (r.json()["id"], r.json()["lastModifiedDateTime"])
            for r in (ecrire(client, f"f{i}.txt", b"x") for i in range(3))
        ]

    premiers = releve()
    assert [t for _, t in premiers] == [
        "2026-01-01T00:01:00Z",
        "2026-01-01T00:02:00Z",
        "2026-01-01T00:03:00Z",
    ]
    client.post("/__admin/drive/reset", headers=ADMIN)
    assert releve() == premiers


def test_supprimer_puis_reecrire_donne_un_NOUVEL_element(client, auth):
    ecrire(client, "reports/report-1.txt", b"hello")
    avant = element(client, auth, "reports/report-1.txt")
    supprimer = "/__admin/drive/files/reports/report-1.txt"
    assert client.delete(supprimer, headers=ADMIN).status_code == 204
    assert client.get(f"/v1.0/drives/{DRIVE}/items/{avant['id']}", headers=auth).status_code == 404
    assert client.delete(supprimer, headers=ADMIN).status_code == 404
    assert ecrire(client, "reports/report-1.txt", b"hello").status_code == 201
    assert element(client, auth, "reports/report-1.txt")["id"] != avant["id"]


@pytest.mark.usefixtures("arbre")
def test_supprimer_un_dossier_emporte_son_contenu(client, auth):
    assert client.delete("/__admin/drive/files/reports", headers=ADMIN).status_code == 204
    assert set(tout_l_arbre(client, auth)) == {"notes.txt"}


@pytest.mark.usefixtures("arbre")
def test_reset_du_lecteur_le_rend_VIDE(client, auth):
    assert client.post("/__admin/drive/reset", headers=ADMIN).status_code == 200
    assert client.get(RACINE, headers=auth).json()["value"] == []


def test_les_noms_invalides_sont_refuses(client):
    ecrire(client, "reports/report-1.txt", b"hello")
    assert ecrire(client, "a:b.txt", b"x").status_code == 400
    assert ecrire(client, "a_vti_b.txt", b"x").status_code == 400
    assert ecrire(client, " espace.txt", b"x").status_code == 400
    assert ecrire(client, "reports", b"x").status_code == 409
    assert ecrire(client, "reports/report-1.txt/inside.txt", b"x").status_code == 409


def test_les_compteurs_ne_retiennent_que_les_appels_servis(client, auth):
    """Ce qu'il faut pour PROUVER qu'un second run idempotent ne retélécharge
    rien : les compteurs, pas le contenu."""
    ecrire(client, "reports/report-1.txt", b"hello")
    client.get(REPORTS)  # 401 : non compté
    item = client.get(f"{REPORTS}?$select=id,name", headers=auth).json()["value"][0]
    telecharger(client, auth, item["id"])
    compteurs = client.get("/__admin/drive/counters", headers=ADMIN).json()
    assert compteurs == {
        "site": 0,
        "drive": 0,
        "drives": 0,
        "item": 0,
        "children": 1,
        "content": 1,
        "download": 1,
    }


def test_le_reset_global_remet_aussi_la_cadence_d_etranglement(client, auth, monkeypatch):
    monkeypatch.setattr(app_module, "THROTTLE_EVERY", 2)
    monkeypatch.setattr(app_module, "_appels", 0)
    assert client.get(SITE, headers=auth).status_code == 200
    assert client.post("/__admin/reset", headers=ADMIN).status_code == 200
    assert client.get(SITE, headers=auth).status_code == 200  # 1er appel après le reset
    assert client.get(SITE, headers=auth).status_code == 429


# ── Le préambule commun : jeton et étranglement ──────────────────────────────

ROUTES_GRAPH = [
    SITE,
    f"/v1.0/sites/{drive_module.SITE_ID}",
    f"/v1.0/sites/{drive_module.SITE_ID}/drive",
    f"/v1.0/sites/{drive_module.SITE_ID}/drives",
    f"/v1.0/drives/{DRIVE}",
    f"/v1.0/drives/{DRIVE}/root",
    RACINE,
    REPORTS,
    f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt",
    f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
    f"/v1.0/drives/{DRIVE}/items/01QUELCONQUE",
    f"/v1.0/drives/{DRIVE}/items/01QUELCONQUE/children",
    f"/v1.0/drives/{DRIVE}/items/01QUELCONQUE/content",
]


@pytest.mark.parametrize("url", ROUTES_GRAPH)
def test_le_Bearer_est_exige_sur_chaque_route(client, url):
    r = client.get(url, follow_redirects=False)
    assert r.status_code == 401
    assert r.json()["error"]["code"] == "InvalidAuthenticationToken"


@pytest.mark.parametrize("url", ROUTES_GRAPH)
def test_l_etranglement_injecte_frappe_chaque_route(client, auth, monkeypatch, url):
    monkeypatch.setattr(app_module, "THROTTLE_EVERY", 1)
    monkeypatch.setattr(app_module, "_appels", 0)
    r = client.get(url, headers=auth, follow_redirects=False)
    assert r.status_code == 429
    assert r.headers["Retry-After"] == app_module.RETRY_AFTER
    assert r.json()["error"]["code"] == "TooManyRequests"


# ── Le contrat ───────────────────────────────────────────────────────────────


def test_le_contrat_publie_est_a_jour_et_couvre_les_fichiers():
    """`make contract` doit avoir été relancé : un client en épingle une copie,
    et un contrat périmé lui ferait valider contre une surface fausse."""
    from entra_mock import app

    publie = yaml.safe_load(Path("contracts/msgraph.openapi.yaml").read_text())
    assert publie == app.openapi()
    chemins = set(publie["paths"])
    assert {
        "/v1.0/sites/{hostname}:/{server_relative_path}",
        "/v1.0/sites/{site_id}/drive",
        "/v1.0/sites/{site_id}/drives",
        "/v1.0/drives/{drive_id}/root:/{chemin}",
        "/v1.0/drives/{drive_id}/items/{item_id}",
        "/v1.0/drives/{drive_id}/items/{item_id}/content",
    } <= chemins
    assert not any(c.startswith(("/__admin", "/sites/")) for c in chemins)
    assert (
        "302"
        in publie["paths"]["/v1.0/drives/{drive_id}/items/{item_id}/content"]["get"]["responses"]
    )


def test_le_plan_de_controle_peut_dater_un_fichier(client, auth):
    r = client.put(
        "/__admin/drive/files/reports/q3.txt?lastModifiedDateTime=2025-10-05T08:00:00Z",
        content=b"q3",
        headers=ADMIN,
    )
    assert r.status_code == 201
    assert element(client, auth, "reports/q3.txt")["lastModifiedDateTime"] == "2025-10-05T08:00:00Z"
    # sans le paramètre, l'horloge du mock reprend la main
    assert ecrire(client, "reports/q3.txt", b"q3 bis").status_code == 200
    assert element(client, auth, "reports/q3.txt")["lastModifiedDateTime"] != "2025-10-05T08:00:00Z"


def test_une_date_mal_ecrite_est_refusee(client):
    r = client.put(
        "/__admin/drive/files/reports/q3.txt?lastModifiedDateTime=05/10/2025",
        content=b"q3",
        headers=ADMIN,
    )
    assert r.status_code == 400 and r.json()["error"]["code"] == "invalidRequest"
