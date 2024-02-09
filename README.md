# xml do xls konvertor

## Popis
Tento program slouží k převodu XML souborů do formátu XLSX.

## Návod
1. Stáhněte si `xml2xls.zip` z [releases]
2. Extrahujte obsah do prázdné složky, kde chcete mít program uložený.
3. Vytvořte složku `xml` a `xls`.
4. Do složky `xml` vložte XML soubory, které chcete převést, do složky `xls se uloží konvertované soubory.
5. Spusťte program `xml2xls.exe`.

> V souboru `results` najdete celkový seznam převedených souborů. s časy lepení a mezi výrobky.
> Vyhodnocování OK/NOK je založena na: 1:30 min. < čas_lepení < 2:00 min. a doba_mezi_výrobky < 1:00 hod., aktuálně nenastavitelné.
> Je důležité, aby nebyla porušena hierarchie složek, jinak program nebude fungovat!!!

### Správná hierarchie:
```
Kořenová složka
└─── xml2xls.exe
│
└─── xml
│   └─── soubor.xml
|
└─── xls
|
└─── templates
│   └─── reclamacion_template.xlsx
│   └─── results_template.xlsx
```