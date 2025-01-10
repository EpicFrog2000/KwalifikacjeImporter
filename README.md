# Instrukcja korzystania

1. **Ustawienie połączenia do bazy danych:**
   - W pliku `Config.json` ustaw odpowiedni ConnectionString do bazy danych w polu `"ConnectionString"`.

2. **Dodawanie folderu z plikami:**
   - Aby dodać folder, w którym będą znajdowały się pliki z kwalifikacjami pracowników, należy utworzyć nowy folder oraz dodać ścieżkę tego folderu do pliku `Config.json`.
   
   **Przykład:**
   
   Mamy folder `C:\Nowe_Pliki`.

   Edytujemy plik `Config.json`:

   **Przed doaniem:**
   ```json
   {
     "Folders_With_Files": [
       "E:\\ITEGER\\Praca\\BhpReader\\Files"
     ],
     "ConnectionString": "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;"
   }
   ```
    **Po dodaniu:**
   ```json
   {
     "Folders_With_Files": [
       "E:\\ITEGER\\Praca\\BhpReader\\Files",
       "C:\\Nowe_Pliki"
     ],
     "ConnectionString": "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;"
   }
   ```
3. Do folderu `C:\Nowe_Pliki` przekopiuj pliki excel z Kwalifikacjami pracowników
4. Uruchomić program KwalifikacjeImporter.exe
5. Po zakończonej pracy programu należy zweryfikować czy w folderze `C:\Nowe_Pliki\Bad_Data` nie zostały utworzone nowe pliki. Jeśli tak to w tych plikach znajdują się dane które nie zostały dodane do bazy danych z powodów opisanych w plikach Errors.txt w folderze `C:\Nowe_Pliki`.