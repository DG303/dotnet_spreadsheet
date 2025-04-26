# 📄 Αυτοματισμός Excel με C# — Χρήση EPPlus

Αυτό το έργο δείχνει πώς να αυτοματοποιήσετε τον χειρισμό αρχείων Excel χρησιμοποιώντας το **EPPlus**, μια ελαφριά, ασφαλή για διακομιστές βιβλιοθήκη για αρχεία `.xlsx` σε C#.

Αντί να βασίζεστε στο Microsoft Office Interop (που απαιτεί εγκατεστημένο το Excel και δεν είναι φιλικό προς τον διακομιστή), το **EPPlus** επιτρέπει την ανάγνωση, εγγραφή και μορφοποίηση αρχείων Excel καθαρά μέσω κώδικα C# — **χωρίς απαίτηση εγκατάστασης Excel**.

## 📋 Προαπαιτούμενα

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download) ή νεότερη έκδοση
- [Visual Studio 2022](https://visualstudio.microsoft.com/) ή [VS Code](https://code.visualstudio.com/) με επεκτάσεις C#
- Βασική κατανόηση της C# και των μορφών αρχείων Excel

## 📦 Τεχνολογίες που χρησιμοποιούνται

- **C# / .NET 8.0**
- **EPPlus 7.0.9** (Πακέτο NuGet)
- **Visual Studio / VS Code**
- **Docker** (για containerization και δοκιμές)

## 📁 Δομή Έργου

```
SpreadsheetEditor/
├── src/
│   ├── SpreadsheetEditor.Core/          # Βασική επιχειρηματική λογική
│   │   └── SpreadsheetEditor.Core.csproj
│   │
│   └── SpreadsheetEditor.Console/       # Εφαρμογή κονσόλας
│       ├── Program.cs                   # Σημείο εισόδου
│       ├── Options.cs                   # Επιλογές γραμμής εντολών
│       └── SpreadsheetEditor.Console.csproj
│
├── data/                                # Αποθήκευση αρχείων Excel (Docker volume)
├── Dockerfile                           # Διαμόρφωση δημιουργίας Docker
├── docker-compose.yml                   # Διαμόρφωση Docker Compose
├── test-docker.sh                       # Σενάριο δοκιμών Docker
└── README.md                           # Τεκμηρίωση έργου
```

## 🚀 Ξεκινώντας

### Κατασκευή και Εκτέλεση Τοπικά

1. Κλωνοποιήστε το αποθετήριο:
```bash
git clone <repository-url>
cd spreadsheet-editor
```

2. Κατασκευάστε τη λύση:
```bash
dotnet build
```

3. Εκτελέστε την εφαρμογή:
```bash
# Δημιουργία νέου αρχείου Excel
dotnet run --project src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj -- --file output.xlsx --new

# Εγγραφή σε κελί
dotnet run --project src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj -- --file output.xlsx --cell A1 --value "Γειά σου, Κόσμε!" --write

# Ανάγνωση από κελί
dotnet run --project src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj -- --file output.xlsx --cell A1 --read
```

### Εκτέλεση με Docker

1. Κατασκευάστε την εικόνα Docker:
```bash
docker-compose build
```

2. Δημιουργήστε έναν κατάλογο για τα αρχεία Excel:
```bash
mkdir -p data
```

3. Εκτελέστε την εφαρμογή:
```bash
# Δημιουργία νέου αρχείου Excel
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --new

# Εγγραφή σε κελί
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --value "Γειά" --write

# Ανάγνωση από κελί
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --read
```

### Εκτέλεση Δοκιμών

1. Εκτελέστε το αυτοματοποιημένο σύνολο δοκιμών Docker:
```bash
chmod +x test-docker.sh  # Κάντε το σενάριο δοκιμών εκτελέσιμο
./test-docker.sh
```

Αυτό το σενάριο δοκιμών θα:
- Δημιουργήσει ένα νέο αρχείο Excel
- Γράψει μια τιμή δοκιμής στο κελί A1
- Διαβάσει την τιμή πίσω και θα επαληθεύσει ότι ταιριάζει
- Καθαρίσει το αρχείο δοκιμής

## 📝 Επιλογές Γραμμής Εντολών

Η εφαρμογή υποστηρίζει τις ακόλουθες επιλογές γραμμής εντολών:

```
-f, --file     Απαιτείται. Διαδρομή προς το αρχείο Excel
-s, --sheet    Προαιρετικό. Όνομα φύλλου εργασίας (προεπιλογή: Sheet1)
-c, --cell     Διεύθυνση κελιού (π.χ., A1), απαιτείται για λειτουργίες ανάγνωσης/εγγραφής
-v, --value    Τιμή για εγγραφή στο κελί
-r, --read     Ανάγνωση τιμής από κελί
-w, --write    Εγγραφή τιμής σε κελί
-n, --new      Δημιουργία νέου αρχείου Excel
--help         Εμφάνιση οθόνης βοήθειας
--version      Εμφάνιση πληροφοριών έκδοσης
```

## ⚙️ Διαμόρφωση

Η εφαρμογή χρησιμοποιεί το `appsettings.json` για διαμόρφωση:

```json
{
  "AppSettings": {
    "DefaultWorksheetName": "Sheet1",
    "DefaultOutputDirectory": "./output"
  }
}
```

## 🔒 Άδεια

Αυτό το έργο χρησιμοποιεί το EPPlus με μη εμπορική άδεια. Για εμπορική χρήση, παρακαλούμε αποκτήστε την κατάλληλη άδεια από το [EPPlus Software](https://epplussoftware.com/).

## 🤝 Συνεισφορά

1. Κάντε fork το αποθετήριο
2. Δημιουργήστε έναν κλάδο χαρακτηριστικών
3. Κάντε commit τις αλλαγές σας
4. Κάντε push στον κλάδο
5. Δημιουργήστε ένα Pull Request

## 📚 Πόροι

- [Τεκμηρίωση EPPlus](https://epplussoftware.com/docs)
- [Τεκμηρίωση .NET](https://docs.microsoft.com/en-us/dotnet/)
- [Τεκμηρίωση Docker](https://docs.docker.com/)

## 🐳 Υποστήριξη Docker

### Προαπαιτούμενα
Εκτός από τα υπάρχοντα προαπαιτούμενα, θα χρειαστείτε:
- [Docker](https://www.docker.com/get-started)
- [Docker Compose](https://docs.docker.com/compose/install/)

### Εκτέλεση με Docker

1. Κατασκευάστε την εικόνα Docker:
```bash
docker-compose build
```

2. Δημιουργήστε έναν κατάλογο για τα αρχεία Excel:
```bash
mkdir -p data
```

3. Εκτελέστε την εφαρμογή:
```bash
# Δημιουργία νέου αρχείου Excel
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --new

# Εγγραφή σε κελί
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --value "Γειά" --write

# Ανάγνωση από κελί
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --read
```

### Εκτέλεση Δοκιμών
Το έργο περιλαμβάνει ένα αυτοματοποιημένο σενάριο δοκιμών για Docker:

```bash
chmod +x test-docker.sh  # Κάντε το σενάριο δοκιμών εκτελέσιμο
./test-docker.sh
```

Αυτό το σενάριο δοκιμών θα:
- Δημιουργήσει ένα νέο αρχείο Excel
- Γράψει μια τιμή δοκιμής στο κελί A1
- Διαβάσει την τιμή πίσω και θα επαληθεύσει ότι ταιριάζει
- Καθαρίσει το αρχείο δοκιμής

### Δομή Έργου Docker
Πρόσθετα αρχεία για υποστήριξη Docker:
```
SpreadsheetEditor/
├── ...
├── data/                                # Αποθήκευση αρχείων Excel (mounted volume)
├── Dockerfile                           # Διαμόρφωση δημιουργίας Docker
├── docker-compose.yml                   # Διαμόρφωση Docker Compose
└── test-docker.sh                       # Σενάριο δοκιμών Docker
``` 