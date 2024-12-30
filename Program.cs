using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Text.Json;

namespace BHPReader
{
    static class ConfigReader
    {
        public class Config_Data_Json
        {
            public string[] Folders_With_Files { get; set; } = { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") };
            public string ConnectionString { get; set; } = string.Empty;
        }
        public static void Get_Config_From_File()
        {
            Try_Init_Json();
            string json = File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json"));
            var config = JsonSerializer.Deserialize<Config_Data_Json>(json);
            if (config != null)
            {
                Program.Folders_With_Files = config.Folders_With_Files;
                Program.ConnectionString = config.ConnectionString;
            }
        }
        private static void Try_Init_Json(){
            if (!File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")))
            {
                File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")).Dispose();
                var defaultConfig = new Config_Data_Json()
                {
                    Folders_With_Files = {},
                    ConnectionString = Program.ConnectionString
                };
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json"), JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
            }
        }
    }
    internal class Program
    {
        public static string[] Folders_With_Files = { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") };
        public static string ConnectionString = "Server=ITEGER-NT;Database=CDN_Wars_Test_4;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
        class BaseInfo
        {
            public int Kod { get; set; } = -1;
            public string Nazwisko { get; set; } = "";
            public string Imie { get; set; } = "";
            public virtual void Load_Info_From_Values(string lp, string kod, string nazwisko, string imie, string? val1, string? val2, string? val3, string? val4) { }
            public virtual void Make_Insert_Command(SqlTransaction transaction, SqlConnection connection, string folder, string file_path) { }
        }
        class Kwalifikacje : BaseInfo
        {
            public string Nazwa_szkolenia { get; set; } = "";
            public DateTime Data_rozpoczecia { get; set; } = DateTime.MinValue;
            public DateTime Data_zakonczenia { get; set; } = DateTime.MinValue;
            public DateTime Data_waznosci { get; set; } = DateTime.MinValue;
            public override void Load_Info_From_Values(string lp, string kod, string nazwisko, string imie, string? val1, string? val2, string? val3, string? val4)
            {
                Kod = int.TryParse(kod, out int kodint) ? kodint : 0;
                Nazwisko = nazwisko;
                Imie = imie;
                Nazwa_szkolenia = val1 ?? "";
                Data_rozpoczecia = DateTime.TryParse(val2, out var dataRozp) ? dataRozp : DateTime.MinValue;
                Data_zakonczenia = DateTime.TryParse(val3, out var dataZak) ? dataZak : DateTime.MinValue;
                Data_waznosci = DateTime.TryParse(val4, out var dataWazn) ? dataWazn : DateTime.MinValue;
            }
            public override void Make_Insert_Command(SqlTransaction transaction, SqlConnection connection, string folder, string file_path)
            {
                var KodPrac = Czy_Istnieje_Pracownik(Kod);
                if (KodPrac != null)
                {
                    string sqlQuery = @"DECLARE @DkmId INT = (select DKM_DkmId from cdn.DaneKadMod where DKM_Nazwa like @Nazwa_szkolenia)
IF @DkmId is null
BEGIN
INSERT INTO [CDN].[DaneKadMod]
           ([DKM_Rodzaj]
           ,[DKM_Nazwa]
           ,[DKM_Opis]
           ,[DKM_Robotnicze]
           ,[DKM_ZjeId]
           ,[DKM_ImportRowId])
     VALUES
           (6
           ,@Nazwa_szkolenia
           ,''
           ,1
           ,''
           ,'');
SET @DkmId = (select DKM_DkmId from cdn.DaneKadMod where DKM_Nazwa like @Nazwa_szkolenia)
END

INSERT INTO [CDN].[Uprawnienia]
           ([UPR_PraId]
		   ,[UPR_DkmId]
           ,[UPR_Rodzaj]
           ,[UPR_Nazwa]
           ,[UPR_WymagLubMowa]
           ,[UPR_UkonLubPismo]
           ,[UPR_StMowa]
           ,[UPR_StPismo]
           ,[UPR_KursOd]
           ,[UPR_KursDo]
           ,[UPR_KursTermin]
           ,[UPR_Opis]
           ,[UPR_Przypomnienie]
           ,[UPR_Zrodlo]
           ,[UPR_Nieaktywne]
           ,[UPR_OpeZalId]
           ,[UPR_StaZalId]
           ,[UPR_TS_Zal]
           ,[UPR_OpeModId]
           ,[UPR_StaModId]
           ,[UPR_TS_Mod]
           ,[UPR_OpeModKod]
           ,[UPR_OpeModNazwisko]
           ,[UPR_OpeZalKod]
           ,[UPR_OpeZalNazwisko])
     VALUES
           (@Kod
		   ,@DkmId
		   ,1
		   ,@Nazwa_szkolenia
           ,0
           ,0
           ,''
           ,''
           ,@Data_rozpoczecia
           ,@Data_zakonczenia
           ,@Data_waznosci
           ,''
           ,1
           ,0
           ,0
           ,1
           ,1
           ,GETDATE()
           ,1
           ,1
           ,GETDATE()
           ,'ADMIN'
           ,'Administrator'
           ,'ADMIN'
           ,'Administrator');";
                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, transaction))
                    {
                        insertCmd.Parameters.AddWithValue("@Kod", KodPrac);
                        insertCmd.Parameters.AddWithValue("@Nazwisko", Nazwisko);
                        insertCmd.Parameters.AddWithValue("@Imie", Imie);
                        insertCmd.Parameters.AddWithValue("@Nazwa_szkolenia", Nazwa_szkolenia);
                        insertCmd.Parameters.AddWithValue("@Data_rozpoczecia", Data_rozpoczecia == DateTime.MinValue ? (object)DBNull.Value : Data_rozpoczecia);
                        insertCmd.Parameters.AddWithValue("@Data_zakonczenia", Data_zakonczenia == DateTime.MinValue ? (object)DBNull.Value : Data_zakonczenia);
                        insertCmd.Parameters.AddWithValue("@Data_waznosci ", Data_waznosci == DateTime.MinValue ? (object)DBNull.Value : Data_waznosci);
                        insertCmd.ExecuteScalar();
                    }
                }
                else
                {
                    throw new Exception($"W bazie nie ma pracownika o akronimie {Kod}, {Nazwisko}, {Imie} z pliku {file_path}");
                }
            }
        }
        class LimityNiobecnosci : BaseInfo
        {
            public int Dni { get; set; } = 0;
            public int Godziny { get; set; } = 0;
            public override void Load_Info_From_Values(string lp, string kod, string nazwisko, string imie, string? val1, string? val2, string? val3, string? val4)
            {
                Kod = int.TryParse(kod, out int kodint) ? kodint : 0;
                Nazwisko = nazwisko;
                Imie = imie;
                Dni = int.TryParse(val1, out var dni) ? dni : 0;
                Godziny = int.TryParse(val2, out var godz) ? godz : 0;
            }
            public override void Make_Insert_Command(SqlTransaction transaction, SqlConnection connection, string cp, string cf)
            {
                return;
                /*string sqlQuery = "";
                using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, transaction))
                {
                    insertCmd.Parameters.AddWithValue("@Kod", Kod);
                    insertCmd.Parameters.AddWithValue("@Nazwisko", Nazwisko);
                    insertCmd.Parameters.AddWithValue("@Imie", Imie);
                    insertCmd.Parameters.AddWithValue("@Dni", Dni);
                    insertCmd.Parameters.AddWithValue("@Godziny", Godziny);
                    insertCmd.ExecuteScalar();
                }*/
            }
        }
        class PPK : BaseInfo
        {
            public DateTime Data_przystapienia { get; set; } = DateTime.MinValue;
            public DateTime Data_wystapienia { get; set; } = DateTime.MinValue;
            public override void Load_Info_From_Values(string lp, string kod, string nazwisko, string imie, string? val1, string? val2, string? val3, string? val4)
            {
                Kod = int.TryParse(kod, out int kodint) ? kodint : 0;
                Nazwisko = nazwisko;
                Imie = imie;
                Data_przystapienia = DateTime.TryParse(val1, out var dataPry) ? dataPry : DateTime.MinValue;
                Data_wystapienia = DateTime.TryParse(val2, out var dataWps) ? dataWps : DateTime.MinValue;
            }
            public override void Make_Insert_Command(SqlTransaction transaction, SqlConnection connection, string cp, string cf)
            {
                return;
                /*string sqlQuery = "";
                using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, transaction))
                {
                    insertCmd.Parameters.AddWithValue("@Kod", Kod);
                    insertCmd.Parameters.AddWithValue("@Nazwisko", Nazwisko);
                    insertCmd.Parameters.AddWithValue("@Imie", Imie);
                    insertCmd.Parameters.AddWithValue("@Data_przystapienia", Data_przystapienia);
                    insertCmd.Parameters.AddWithValue("@Data_wystapienia", Data_wystapienia);
                    insertCmd.ExecuteScalar();
                }*/
            }
        }
        private static void Usun_Ukryte_Karty(XLWorkbook workbook)
        {
            var hiddenSheets = new List<IXLWorksheet>();
            foreach (var sheet in workbook.Worksheets)
            {
                if (sheet.Visibility == XLWorksheetVisibility.Hidden)
                {
                    hiddenSheets.Add(sheet);
                }
            }
            foreach (var sheet in hiddenSheets)
            {
                workbook.Worksheets.Delete(sheet.Name);
            }
            workbook.Save();
        }
        public static int Main()
        {
            string current_folder = "";
            string current_file_path = "";
            try
            {
                ConfigReader.Get_Config_From_File();
                foreach (var folder in Folders_With_Files)
                {
                    current_folder = folder;
                    var ListyList = new List<List<BaseInfo>>();
                    foreach (var path in Directory.GetFiles(folder))
                    {
                        current_file_path = path;
                        if (!path.Contains(".txt"))
                        {
                            Console.WriteLine($"Czytanie pliku {path}");
                            XLWorkbook workbook = new XLWorkbook(path);
                            Usun_Ukryte_Karty(workbook);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var listadanych = ReadZakladka(worksheet, Get_Typ_Zakladki(worksheet));
                                if (listadanych.Count > 0)
                                {
                                    ListyList.Add(listadanych);
                                }
                            }
                        }
                            MoveProcessedFile(current_file_path, current_folder);
                    }
                    if (ListyList.Count > 0)
                    {
                        Console.WriteLine($"Insering data to db ...");
                        var NieWpisaneDane = Insert_To_Db(ListyList, current_folder, current_file_path);
                        Save_Bad_Data(NieWpisaneDane, current_folder);
                    }
                }
            }
            catch (Exception ex)
            {
                SaveErrorToFile(ex.Message, current_folder, current_file_path);
                Console.WriteLine(ex.Message);
            }
            return 0;
        }
        private static void Save_Bad_Data(List<BaseInfo> data, string outputPath)
        {
            if (data == null || !data.Any())
            {
                return;
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Kwalifikacje");
                worksheet.Cell(1, 1).Value = "Lp";
                worksheet.Cell(1, 2).Value = "Kod";
                worksheet.Cell(1, 3).Value = "Nazwisko";
                worksheet.Cell(1, 4).Value = "Imię";
                worksheet.Cell(1, 5).Value = "Nazwa szkolenia";
                worksheet.Cell(1, 6).Value = "Data rozpoczęcia";
                worksheet.Cell(1, 7).Value = "Data zakończenia";
                worksheet.Cell(1, 8).Value = "Data ważności";

                for (int i = 1; i <= data.Count; i++)
                {
                    worksheet.Cell(1 + i, 1).Value = i;
                    worksheet.Cell(1 + i, 2).Value = data[i - 1].Kod;
                    worksheet.Cell(1 + i, 3).Value = data[i - 1].Nazwisko;
                    worksheet.Cell(1 + i, 4).Value = data[i - 1].Imie;

                    var kwalifikacje = data[i - 1] as Kwalifikacje;
                    if (kwalifikacje != null)
                    {
                        worksheet.Cell(1 + i, 5).Value = kwalifikacje.Nazwa_szkolenia;
                        if (kwalifikacje.Data_rozpoczecia != DateTime.MinValue)
                        {
                            worksheet.Cell(1 + i, 6).Value = kwalifikacje.Data_rozpoczecia;
                        }
                        if (kwalifikacje.Data_zakonczenia != DateTime.MinValue)
                        {
                            worksheet.Cell(1 + i, 7).Value = kwalifikacje.Data_zakonczenia;
                        }
                        if (kwalifikacje.Data_waznosci != DateTime.MinValue)
                        {
                            worksheet.Cell(1 + i, 8).Value = kwalifikacje.Data_waznosci;
                        }
                    }
                }

                try
                {
                    if(!Directory.Exists(outputPath + "\\Bad_Data"))
                    {
                        Directory.CreateDirectory(outputPath + "\\Bad_Data");
                    }
                    workbook.SaveAs(outputPath + "\\Bad_Data\\Bad_Data.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Wystąpił błąd podczas zapisywania pliku z błędnymi danymi: {ex.Message}");
                    throw new Exception($"Wystąpił błąd podczas zapisywania pliku z błędnymi danymi: {ex.Message}");
                }
            }
        }
        private static int Get_Typ_Zakladki(IXLWorksheet worksheet)
        {
            var val5 = worksheet.Cell(1, 5).GetFormattedString().Trim().ToLower();
            if (val5.Contains("data_przystapienia"))
            {
                return 1;
            }
            else if (val5.Contains("dni"))
            {
                return 2;
            }
            else if (val5.Contains("nazwa_szkolenia"))
            {
                return 3;
            }
            return 0;
        }
        private static List<BaseInfo> ReadZakladka(IXLWorksheet worksheet, int Typ)
        {
            var result = new List<BaseInfo>();
            int i = 2;
            BaseInfo Init()
            {
                if (Typ == 1)
                {
                    return new PPK();
                }else if (Typ == 2)
                {
                    return new LimityNiobecnosci();
                }else if (Typ == 3)
                {
                    return new Kwalifikacje();
                }
                throw new Exception("Nieznany typ pliku");
            }

            if (Typ == 3)
            {
                while (true)
                {
                    if (worksheet.Row(i).IsHidden == false)
                    {
                        var Lp = worksheet.Cell(i, 1).GetFormattedString().Trim();
                        if (string.IsNullOrEmpty(Lp))
                        {
                            break;
                        }
                        var Kod = worksheet.Cell(i, 2).GetFormattedString().Trim();
                        var Nazwisko = worksheet.Cell(i, 3).GetFormattedString().Trim();
                        var Imie = worksheet.Cell(i, 4).GetFormattedString().Trim();
                        var val1 = worksheet.Cell(i, 5).GetFormattedString().Trim();
                        var val2 = worksheet.Cell(i, 6).GetFormattedString().Trim();
                        var val3 = worksheet.Cell(i, 7).GetFormattedString().Trim();
                        var val4 = worksheet.Cell(i, 8).GetFormattedString().Trim();
                        var obj = Init();
                        obj.Load_Info_From_Values(Lp, Kod, Nazwisko, Imie, val1, val2, val3, val4);
                        result.Add(obj);
                    }
                    i++;
                }
            }
            return result;
        }
        private static List<BaseInfo> Insert_To_Db(List<List<BaseInfo>> Objects, string folder, string file_path)
        {
            var niewpisanedane = new List<BaseInfo>();
            foreach (var ObjectList in Objects)
            {
                foreach (var Data in ObjectList)
                {
                    try
                    {
                        using (var connection = new SqlConnection(ConnectionString))
                        {
                            connection.Open();
                            using (var transaction = connection.BeginTransaction())
                            {
                                Data.Make_Insert_Command(transaction, connection, folder, file_path);
                                transaction.Commit();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        niewpisanedane.Add(Data);
                        Console.WriteLine(ex.Message);
                        SaveErrorToFile(ex.Message, folder, file_path);
                    }
                }
            }
            Console.WriteLine("Poprawnie dodano dane do bazy");
            return niewpisanedane;
        }
        private static int? Czy_Istnieje_Pracownik(int Kod)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    string query = "SELECT PRI_PraId FROM cdn.Pracidx WHERE PRI_Kod = @Kod and PRI_Typ = 1;";
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@Kod", Kod);
                        var result = cmd.ExecuteScalar();
                        if (result == DBNull.Value || result == null)
                        {
                            return null;
                        }
                        return Convert.ToInt32(result);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Błąd: " + ex.Message);
                throw;
            }
        }
        private static void SaveErrorToFile(string Value, string savePath, string nazwaPlkiu ) {
            try {
                var filePath = Path.Combine(savePath, "Errors.txt");
                if (!File.Exists(filePath))
                {
                    var fs = File.Create(filePath);
                    fs.Dispose();
                }
                File.AppendAllText(filePath, Value + Environment.NewLine);
            } catch (Exception ex)
            {
                Console.WriteLine("Błąd zapisywania errora do pliku: " + ex.Message);
            }
        }
        private static void MoveProcessedFile(string filepath, string folder)
        {
            string processedFolder = Path.Combine(folder, "Processed_Files");
            if (!Directory.Exists(processedFolder))
            {
                Directory.CreateDirectory(processedFolder);
            }
            string destinationPath = Path.Combine(processedFolder, Path.GetFileName(filepath));
            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }
            File.Move(filepath, destinationPath);
        }
    }
}