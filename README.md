# Circet Mail Generator

Web app + complément Outlook (moderne) pour générer rapidement des emails avec :
- Import Excel (_Type Mail, Description Mail, Destinataires, Copie_) → remplit "À" et "Cc" selon **Type + Région**
- Mapping Type/Région → corps par défaut (sauvegardé en localStorage)
- Génération Objet + Corps (tokens `{REGION}`, `{DATE}`, `{HEURE}`, `{SEMAINE}`)
- **Pièce jointe locale** (depuis le complément Outlook)

## Déploiement GitHub Pages

1. Crée le repo **circet-mail-generator** et colle ce dossier.
2. Paramètres → Pages → Branch: `main` / Root.
3. Remplace `daveisa88` dans `addin/manifest.xml` si ton identifiant GitHub est différent.
4. Optionnel: ajoute un `assets/logo.png` (et mets à jour l’URL dans le manifest si tu changes le chemin).

## Sideload (installation) du complément

### Outlook Web
- ⚙️ → **Gérer les compléments** → **Mes compléments** → **Téléverser à partir d’un fichier**  
- Sélectionne `addin/manifest.xml` (tu peux le télécharger depuis le repo).

### Outlook Bureau
- Accueil → **Obtenir des compléments** → **Mes compléments** → **Ajouter un complément personnalisé** → **Ajouter à partir d’un fichier**.

> Le complément apparaît automatiquement **lorsque tu crées un nouveau message** (form type *Compose*). Il ouvre le volet avec l’application web.  
> Utilise le bouton **“Insérer dans Outlook (complément)”** pour pousser À/Cc/Objet/Corps + **ajouter la PJ**.

## Excel attendu

Première feuille (haut de tableau en entêtes) :

```
_Type Mail | Description Mail | Destinataires | Copie
```

- La correspondance cherche `Description Mail` == `"<Type> <Région>"` (ex : `Avancement de prod IDF`).
- Les emails peuvent être séparés par `;` ou `,`.

## Limites / Notes

- Le bouton **Ouvrir Outlook (mailto)** conserve la signature Outlook, mais **ne peut pas** attacher un fichier (limitation mailto).  
- Pour joindre un fichier local automatiquement, utilise **le complément** (bouton “Insérer dans Outlook (complément)”).  
- L’attache base64 nécessite Mailbox **1.8+** (Outlook Microsoft 365 actuel ✅).
