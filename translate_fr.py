# TODO
import re
from pathlib import Path
from pprint import pprint


def word_list(filename):
    words = set()
    with Path(filename).open("r", encoding="utf-8") as s:
        for line in s:
            if line.startswith("REM") or line.startswith("'"):
                continue

            words.update(re.findall(r"\W(\w+)\W", line))
    return sorted(words)


def translate(filename, d):
    lines = []
    with Path(filename).open("r", encoding="utf-8") as s:
        for line in s:
            line = line.rstrip()
            for k, v in d.items():
                if v is not None:
                    line = re.sub(r"(^|\W){}(\W|$)".format(k), r"\1{}\2".format(v), line)
            lines.append(line)

    return "\n".join(lines)


d = {
    'AddSetElement': 'AjouterElementAEnsemble',
    'AddSetElements': 'AjouterElementsAEnsemble',
    'AppendListElement': 'AjouterElementAListe',
    'AppendListElements': 'AjouterElementsAListe',
    'ArrayList': 'Liste',
    'ArrayToString': 'TableauVersChaine',
    'Contains': 'Contient',
    'CopyArray': 'CopierTableau',
    'CopyMap': 'CopierDict',
    'EnumToArray': 'EnumVersTableau',
    'GetEnumSize': 'TrouverTailleEnum',
    'GetListElement': 'ObtenirElementListe',
    'GetListSize': 'TrouverTailleListe',
    'GetMapElement': 'ObtenirElementDict',
    'GetMapElementOrDefault': 'ObtenirElementDictOuDefaut',
    'GetMapSize': 'TrouverTailleDict',
    'GetSetSize': 'TrouverTailleEnsemble',
    'HashMap': 'Dict',
    'HashSet': 'Ensemble',
    'InsertListElement': 'InsererElementDansListe',
    'ListIndexOf': 'TrouverIndexDansListe',
    'ListIsEmpty': 'ListeEstVide',
    'ListLastIndexOf': 'TrouverDernierIndexDansListe',
    'ListToArray': 'ListeVersTableau',
    'ListToString': 'ListeVersChaine',
    'MapContains': 'DictContient',
    'MapIsEmpty': 'DictEstVide',
    'MapKeysToArray': 'DictClesVersTableau',
    'MapKeysToSet': 'DictClesVersEnsemble',
    'MapToString': 'DictVersChaine',
    'MapValuesToArray': 'DictValeursVersTableau',
    'MergeMaps': 'FusionneDicts',
    'NewEmptyMap': 'CreerDictVide',
    'NewEmptySet': 'CreerEnsembleVide',
    'NewList': 'CreerListe',
    'NewListFromArray': 'CreerListeDepuisTableau',
    'NewListWithCapacity': 'CreerListeAvecCapacite',
    'NewSet': 'CreerEnsemble',
    'NewSetFromArray': 'CreerEnsembleDepuisTableau',
    'PopListElement': 'EnleverElementDeListe',
    'PutMapElement': 'AjouterElementADict',
    'RemoveListElement': 'EnleverElementDeListe',
    'RemoveMapElement': 'EnleverElementDeDict',
    'RemoveSetElement': 'EnleverElementDeEnsemble',
    'ReverseList': 'RenverserListe',
    'ReversedArray': 'TableauRenverse',
    'SetContains': 'EnsembleContient',
    'SetIsEmpty': 'EnsembleEstVide',
    'SetListCapacity': 'ChangerCapaciteListe',
    'SetListElement': 'MettreElementDansListe',
    'SetToArray': 'EnsembleVersTableau',
    'SetToString': 'EnsembleVersChaine',
    'ShuffleList': 'MelangerListe',
    'ShuffledArray': 'TableauMelange',
    'SortArrayInPlace': 'TrierTableauEnPlace',
    'SortList': 'TrierListe',
    'SortedArray': 'TableauTrie',
    'TakeSetElement': 'PrendreElementDeEnsemble',
    'UnoValueToString': 'ValeurUnoVersChaine',
    '_CopySwap': '_CopierEtPermuter',
    '_EnsureListCapacity': '_GarantirCapaciteDeListe',
    '_Merge': '_Fusionner',
    '_Partition': '_Partitionner',
    '_QuickSort': '_TriRapide',
    'arr': 'tab',
    'arr1': 'tab1',
    'arr2': 'tab2',
    'capacity': "capacite",
    'cur': 'cour',
    'default': 'defaut',
    'elementsSize': 'tailleElements',
    'error': 'erreur',
    'flush': 'purge',
    'initialCapacityOrArr': 'capaciteInitialeOuTableau',
    'key': 'cle',
    'keyTypeName': 'nomTypeCle',
    'list': 'liste',
    'm1': 'd1',
    'm2': 'd2',
    'map': 'dict',
    'newArr': 'nouveauTableau',
    'newCapacity': 'nouvelleCapacite',
    'newM': 'nouveauD',
    'other': 'autre',
    'remove': 'enleve',
    'reversed': 'renverse',
    'size': 'taille',
    'swap': 'permute',
    'value': 'valeur',
    'valueTypeName': 'nomTypeValeur',
    'm': 'd',
}

if __name__ == "__main__":
    # pprint(dict.fromkeys(word_list("Collections.vb")))
    print(translate("Collections.vb", d))
    print(translate("TestCollections.vb", d))
