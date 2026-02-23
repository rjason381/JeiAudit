# -*- coding: utf-8 -*-
"""
Verifica existencia de parametros por nombre y categoria en Revit,
independientemente de si tienen valor.

Ejecucion esperada:
- pyRevit (panel Python Script)
- RevitPythonShell
"""

import os
import csv
import unicodedata
from collections import defaultdict

import clr  # type: ignore
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (  # type: ignore
    BuiltInCategory,
    Category,
    InstanceBinding,
    TypeBinding,
)


def normalize(text):
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )
    return text


def get_doc():
    try:
        return __revit__.ActiveUIDocument.Document  # pyRevit/RevitPythonShell
    except Exception:
        raise Exception(
            "No se detecto contexto de Revit. Ejecuta este script dentro de Revit (pyRevit o RevitPythonShell)."
        )


def category_id(doc, bic):
    cat = Category.GetCategory(doc, bic)
    return cat.Id.IntegerValue if cat else None


def collect_bindings(doc):
    bindings = defaultdict(list)
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()

    while it.MoveNext():
        definition = it.Key
        binding = it.Current

        name = definition.Name
        try:
            cats = set(c.Id.IntegerValue for c in binding.Categories)
        except Exception:
            cats = set()

        if isinstance(binding, InstanceBinding):
            kind = "Instance"
        elif isinstance(binding, TypeBinding):
            kind = "Type"
        else:
            kind = type(binding).__name__

        guid_text = ""
        try:
            if hasattr(definition, "GUID") and definition.GUID:
                guid_text = str(definition.GUID)
        except Exception:
            pass

        bindings[name].append(
            {
                "kind": kind,
                "cat_ids": cats,
                "guid": guid_text,
            }
        )

    return bindings


def build_required_matrix(doc):
    # Categorias (ES)
    cats = {
        "Muros": category_id(doc, BuiltInCategory.OST_Walls),
        "Puertas": category_id(doc, BuiltInCategory.OST_Doors),
        "Ventanas": category_id(doc, BuiltInCategory.OST_Windows),
        "Suelos": category_id(doc, BuiltInCategory.OST_Floors),
        "Habitaciones": category_id(doc, BuiltInCategory.OST_Rooms),
    }

    required = {
        "ARG_SECTOR": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_NIVEL": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_SISTEMA": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_UNIFORMAT": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_CODIGO DE PARTIDA": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_UNIDAD DE PARTIDA": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
        "ARG_DESCRIPCION DE PARTIDA": ["Muros", "Puertas", "Ventanas", "Suelos", "Habitaciones"],
    }

    return required, cats


def find_similar_names(target_name, all_names):
    target_norm = normalize(target_name)
    hits = []
    for name in all_names:
        if normalize(name) == target_norm:
            hits.append(name)
    return hits


def main():
    doc = get_doc()
    bindings = collect_bindings(doc)
    required, categories = build_required_matrix(doc)

    out_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    out_csv = os.path.join(out_dir, "reporte_existencia_parametros_peb.csv")

    rows = []
    missing = 0

    all_binding_names = list(bindings.keys())

    for param_name, cat_names in required.items():
        defs = bindings.get(param_name, [])
        exists_by_name = len(defs) > 0

        exact_or_accent_variants = find_similar_names(param_name, all_binding_names)

        for cat_name in cat_names:
            cat_id = categories.get(cat_name)
            bound_to_category = False
            kinds = set()
            guids = set()

            for d in defs:
                kinds.add(d["kind"])
                if d["guid"]:
                    guids.add(d["guid"])
                if cat_id is not None and cat_id in d["cat_ids"]:
                    bound_to_category = True

            status = "OK" if (exists_by_name and bound_to_category) else "FALTA"
            if status == "FALTA":
                missing += 1

            rows.append(
                {
                    "Parametro": param_name,
                    "Categoria": cat_name,
                    "ExisteNombreExacto": "SI" if exists_by_name else "NO",
                    "VinculadoCategoria": "SI" if bound_to_category else "NO",
                    "TipoVinculacion": "/".join(sorted(kinds)) if kinds else "-",
                    "GUIDs": ";".join(sorted(guids)) if guids else "-",
                    "CoincidenciasAcentoMayus": ";".join(sorted(exact_or_accent_variants)) if exact_or_accent_variants else "-",
                    "Estado": status,
                }
            )

    fieldnames = [
        "Parametro",
        "Categoria",
        "ExisteNombreExacto",
        "VinculadoCategoria",
        "TipoVinculacion",
        "GUIDs",
        "CoincidenciasAcentoMayus",
        "Estado",
    ]

    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    total = len(rows)
    ok = total - missing
    print("==== REPORTE EXISTENCIA PARAMETROS PEB ====")
    print("Archivo:", out_csv)
    print("Total verificaciones:", total)
    print("OK:", ok)
    print("FALTANTES:", missing)


if __name__ == "__main__":
    main()
