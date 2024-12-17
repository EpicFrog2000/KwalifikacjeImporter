using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Text.Json;

namespace BHPReader
{
    static class ConfigReader{
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
        abstract class BaseInfo
        {
            public int Kod { get; set; } = 0;
            public string Nazwisko { get; set; } = "";
            public string Imie { get; set; } = "";
            public abstract void Load_Info_From_Values(string lp, string kod, string nazwisko, string imie, string? val1, string? val2, string? val3, string? val4);
            public abstract void Make_Insert_Command(int KodPrac, SqlTransaction transaction, SqlConnection connection);
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
            public override void Make_Insert_Command(int KodPrac, SqlTransaction transaction, SqlConnection connection)
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
            public override void Make_Insert_Command(int KodPrac, SqlTransaction transaction, SqlConnection connection)
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
            public override void Make_Insert_Command(int KodPrac, SqlTransaction transaction, SqlConnection connection)
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
        public static int Main()
        {
            string cf = "";
            string cp = "";
            try
            {
                ConfigReader.Get_Config_From_File();
                foreach (var folder in Folders_With_Files)
                {
                    cf = folder;
                    var ListyList = new List<List<BaseInfo>>();
                    foreach (var path in Directory.GetFiles(folder))
                    {
                        cp = path;
                        if (!path.Contains(".txt"))
                        {
                            Console.WriteLine($"Czyanie pliku {path}");
                            IXLWorkbook workbook = new XLWorkbook(path);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var listadanych = ReadZakladka(worksheet, Get_Typ_Zakladki(worksheet));
                                if (listadanych.Count > 0)
                                {
                                    ListyList.Add(listadanych);
                                }
                            }
                        }
                    }
                    if (ListyList.Count > 0)
                    {
                        Console.WriteLine($"Insering data to db ...");
                        Insert_To_Db(ListyList, cf, cp);
                    }
                }
            }
            catch (Exception ex)
            {
                SaveErrorToFile(ex.Message, cf, cp);
                Console.WriteLine(ex.Message);
            }
            return 0;
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
        private static List<BaseInfo>  ReadZakladka(IXLWorksheet worksheet, int Typ)
        {
            BaseInfo CreateInstance()
            {
                return Typ switch
                {
                    1 => new PPK(),
                    2 => new LimityNiobecnosci(),
                    3 => new Kwalifikacje(),
                    _ => throw new ArgumentException("Nierozpoznany typ zakładki")
                };
            }
            var result = new List<BaseInfo>();
            int i = 2;
            while (true)
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
                var obj = CreateInstance();
                obj.Load_Info_From_Values(Lp, Kod, Nazwisko, Imie, val1, val2, val3, val4);
                result.Add(obj);
                i++;
            }
            return result;
        }
        private static void Insert_To_Db(List<List<BaseInfo>> Objects,string cp, string cf)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();
                foreach (var ObjectList in Objects)
                {
                    foreach (var Data in ObjectList)
                    {
                        try
                        {
                            var KodPrac = Czy_Istnieje_Pracownik(Data.Kod);
                            if (KodPrac != null)
                            {
                                Data.Make_Insert_Command((int)KodPrac, transaction, connection);
                            }
                            else
                            {
                                SaveErrorToFile($"W bazie nie ma pracownika o akronimie {Data.Kod}, {Data.Nazwisko}, {Data.Imie} w pliku {cf}", cf, cp);
                                Console.WriteLine($"W bazie nie ma pracownika o akronimie {Data.Kod}, {Data.Nazwisko}, {Data.Imie}");
                            }
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            throw;
                        }
                    }
                }
                transaction.Commit();
                connection.Close();
                Console.WriteLine("Poprawnie dodano dane do bazy");
            }
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
        private static void SaveErrorToFile(string Value, string nazwaPlkiu, string savePath){
            try{
                var filePath = Path.Combine(savePath, "Errors.txt");
                if (!File.Exists(filePath))
                {
                    var fs = File.Create(filePath);
                    fs.Dispose();
                }
                File.AppendAllText(filePath, Value + " W pliku: " + nazwaPlkiu  + Environment.NewLine);
            }catch(Exception ex)
            {
                Console.WriteLine("Błąd zapisywania errora do pliku: " + ex.Message);
            }
        }
    }
}