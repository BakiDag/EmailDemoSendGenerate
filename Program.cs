using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Net.Mail;
using FluentEmail.Core; //mit nuget installiert fluentmail smtp
using FluentEmail.Smtp;
using System.Diagnostics;

// Bitte schreiben Sie eine Konsolenanwendung, die für einen angegebenen Namen eine E-Mail-Adresse nach folgendem Muster erzeugt und ausgibt:

//Lisa Shell –          a.lische@company.net 
//Lisa Shell -          a.lishel@company.net 
//Frank Ebersoll –      k.freber@company.net
//Ali Tas –             i.altas @company.net
//Gernot von Hagen –    t.gehage@company.net
//Julia Wehmeier –      a.juwehm@company.net 
//Henrik Gerner –       k.hegern@company.net 

//Namen sollen sowohl über die Kommandozeile als auch aus einer Textdatei (separiert durch Zeilenschaltung) eingelesen werden können.

namespace EmailProgramm
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string Pfad = @"C:\EmailProgram\";
            string Dateiname = "EmailAdressBuch.txt";
            string PathTOTXTFile = Pfad + Dateiname;
            string outputfile = "outputfile.txt";
            string OutputFileFromTXT = Pfad + outputfile;

            string CSVDteiname = "EmailAdressBuchCSV.csv";
            string PathToCSVFile = Pfad + CSVDteiname;
            string EnteredName = null;
            string[] SmallArray4Names;
            string completeMailAdress = null;

            string AtCompany = "@company.net";
            string[,] Inhalt = new string[7, 4];

            try // Pfadpruefung
            {
                if (Directory.Exists(Pfad))
                {
                    Console.WriteLine("Der Pfad existiert: " + Pfad);
                    //Console.WriteLine("Ordner wird geoeffnet");
                    //Process.Start(Pfad);

                }
                else
                {
                    DirectoryInfo OrdnerAufC = Directory.CreateDirectory(Pfad);
                    Console.WriteLine("Der Ordner wurde erfolgreich erstellt um {0}", Directory.GetCreationTime(Pfad) + " im Verzeichnis: " + Pfad);
                    //Process.Start(Pfad);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei der Durchfuehrung: {0}", e.ToString());
            }                                                                                                           //Pfadpruefung

            try//Filepfuefung und Inhalt schreiben
            {
                if (!File.Exists(PathTOTXTFile))
                {

                    Console.WriteLine("Datei wird erstellt");

                    using (StreamWriter TXTDatei = File.CreateText(PathTOTXTFile))
                    {
                        Console.WriteLine("Inhalt wird geschrieben");
                        TXTDatei.WriteLine("Lisa Shell");
                        TXTDatei.WriteLine("Frank Ebersoll");
                        TXTDatei.WriteLine("Ali Tas");
                        TXTDatei.WriteLine("Gernot von Hagen");
                        TXTDatei.WriteLine("Julia Wehmeier");
                        TXTDatei.WriteLine("Henrik Gerner");
                        Thread.Sleep(5000);
                        Console.Clear();
                    }
                }
                else
                {
                    Console.WriteLine("Datei existiert: " + PathTOTXTFile);
                    Thread.Sleep(5000);
                    Console.Clear();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei der Durchfuehrung: {0}", e.ToString());
            }                                                                                                           //Filepfuefung und Inhalt in TXT Datei schreiben

            try//Inhalt TXT Datei ausgeben
            {
                using (StreamReader DeineDatei = File.OpenText(PathTOTXTFile))
                {
                    string Inhalt2;
                    Console.WriteLine("Inhalt der Textdatei wird ausgegeben");
                    while ((Inhalt2 = DeineDatei.ReadLine()) != null)
                    {
                        Console.WriteLine(Inhalt2);
                    }
                }
                Thread.Sleep(5000);
                Console.Clear();
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei der Durchfuehrung: {0}", e.ToString());
            }                                                                                                           //Inhalt TXT Datei ausgeben
            try // Textdatei laden, Mail Adressen generieren, Name + Mail exportieren
            {
                using (StreamReader Textdatei = File.OpenText(PathTOTXTFile))
                {
                    var Zeilen = File.ReadAllLines(PathTOTXTFile);
                    if (!File.Exists(OutputFileFromTXT))
                    {
                        Console.WriteLine("Datei wird erstellt {0}", OutputFileFromTXT);
                        Thread.Sleep(5000);
                        Console.Clear();

                    }
                    else
                    {
                        Console.WriteLine("Datei existiert: " + OutputFileFromTXT);
                        Thread.Sleep(5000);
                        Console.Clear();
                    }
                    using (FileStream t = new FileStream(OutputFileFromTXT, FileMode.Append, FileAccess.Write))
                    using (StreamWriter ausgabe = new StreamWriter(t))
                    {
                        ausgabe.WriteLine("Name;Email;");
                    }//Spalten Ueberschrift fuer Excel falls gewuenscht
                    for (int i = 0; i < Zeilen.Length; i++)
                    {
                        Inhalt[i, 0] += Textdatei.ReadLine();

                        using (FileStream t = new FileStream(OutputFileFromTXT, FileMode.Append, FileAccess.Write))
                        using (StreamWriter ausgabe = new StreamWriter(t))
                        {
                            ausgabe.Write(Inhalt[i, 0] + ";");
                            Console.WriteLine("Name {0}  wird in {1} exportiert", Inhalt[i, 0], OutputFileFromTXT);
                            var x = (Inhalt[i, 0].Split(' '));
                            if (x[1].Contains(' '))
                            {
                                x[2] = x[1].Split(' ').ToString();
                                Console.WriteLine("if");
                                ausgabe.WriteLine(x[2]);
                            }

                            if (x.Length <= 2)
                            {
                                if (x[1].Length < 4)
                                {
                                    completeMailAdress = x[0].Last() + "." + x[0].Substring(0, 2) + x[1] + AtCompany;
                                    Console.WriteLine("Email Adresse aus dem Namen generiert");
                                    ausgabe.WriteLine(completeMailAdress.ToLower());
                                    Console.WriteLine("email {0}  wird in {1} exportiert", completeMailAdress.ToLower(), OutputFileFromTXT);
                                }
                                else
                                {
                                    completeMailAdress = x[0].Last() + "." + x[0].Substring(0, 2) + x[1].Substring(0, 4) + AtCompany;
                                    Console.WriteLine("Email Adresse aus dem Namen generiert");
                                    ausgabe.WriteLine(completeMailAdress.ToLower());
                                    Console.WriteLine("email {0}  wird in {1} exportiert", completeMailAdress.ToLower(), OutputFileFromTXT);
                                }
                            }
                            else
                            {
                                if (x[x.Length - 1].Length < 4)
                                {
                                    completeMailAdress = x[0].Last() + "." + x[0].Substring(0, 2) + x[x.Length - 1] + AtCompany;
                                    Console.WriteLine("Email Adresse aus dem Namen generiert");
                                    ausgabe.WriteLine(completeMailAdress.ToLower());
                                    Console.WriteLine("email {0}  wird in {1} exportiert", completeMailAdress.ToLower(), OutputFileFromTXT);
                                }
                                else
                                {
                                    completeMailAdress = x[0].Last() + "." + x[0].Substring(0, 2) + x[x.Length - 1].Substring(0, 4) + AtCompany;
                                    Console.WriteLine("Email Adresse aus dem Namen generiert");
                                    ausgabe.WriteLine(completeMailAdress.ToLower());
                                    Console.WriteLine("email {0}  wird in {1} exportiert", completeMailAdress.ToLower(), OutputFileFromTXT);
                                }
                            }
                        }
                    }
                    Thread.Sleep(3000);
                    Console.Clear();
                }


            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler: {0}", e.ToString());
            }                                                                                                           // Textdatei laden, Mail Adressen generieren, Name + Mail exportieren
            int loop = 0;
            while (loop == 0)// Dateiendung andern Abfrage
            {
                Console.WriteLine("Moechten Sie die Datei auf CSV aendern um sie mit Excel zu bearbeiten? ja/nein");
                string Antwort = Console.ReadLine();
                switch (Antwort.ToLower())
                {
                    case "ja":
                        try
                        {
                            FileInfo f = new FileInfo(OutputFileFromTXT);
                            if (File.Exists(Pfad + "outputfile.csv"))
                            {
                                Console.WriteLine("Eine Datei mit dem selben Namen existiert bereits.");
                                Console.WriteLine("Vorhandene Datei wird gesichert.");
                                var newfilename = DateTime.Now;
                                //var newfilename = File.GetCreationTime(Pfad + "outputfile.csv");            // Throws exception if file exists with the same creation time                    
                                File.Move(Pfad + "outputfile.csv", Pfad + newfilename.ToString("ddMMyyyyHHmmss") + ".csv");


                            }
                            f.MoveTo(Path.ChangeExtension(OutputFileFromTXT, ".csv"));
                            Console.WriteLine("Dateiendung geaendert");
                            Thread.Sleep(3000);
                            Console.Clear();
                            loop = 1;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Fehler bei der Dateierstellung: {0}", e.ToString());
                        }
                        break;
                    case "nein":
                        Console.WriteLine("ok");
                        loop = 1;
                        break;
                    default:
                        Console.WriteLine("Falsche Eingabe.");
                        break;
                }
            }                                                                                             // Dateiendung andern Abfrage
            Console.WriteLine("Manuelles generieren der Email Adresse:");
            Console.WriteLine();
            Console.WriteLine("Geben Sie hierzu einen Vornamen und einen Nachnamen ein, aus dem eine Adresse generiert wird:");
            EnteredName = Console.ReadLine();
            SmallArray4Names = EnteredName.Split(' ');

            try //Email Adresse generieren aus der manuellen Eingabe
            {
                if (SmallArray4Names.Length <= 2)
                {

                    if (SmallArray4Names[1].Length < 4)
                    {
                        completeMailAdress = SmallArray4Names[0].Last() + "." + SmallArray4Names[0].Substring(0, 2) + SmallArray4Names[1] + AtCompany;
                        Console.WriteLine(completeMailAdress.ToLower());
                    }
                    else
                    {
                        completeMailAdress = SmallArray4Names[0].Last() + "." + SmallArray4Names[0].Substring(0, 2) + SmallArray4Names[1].Substring(0, 4) + AtCompany;
                        Console.WriteLine(completeMailAdress.ToLower());
                    }
                }
                else
                {
                    if (SmallArray4Names[SmallArray4Names.Length - 1].Length < 4)
                    {
                        completeMailAdress = SmallArray4Names[0].Last() + "." + SmallArray4Names[0].Substring(0, 2) + SmallArray4Names[SmallArray4Names.Length - 1] + AtCompany;
                        Console.WriteLine(completeMailAdress.ToLower());
                    }
                    else
                    {
                        completeMailAdress = SmallArray4Names[0].Last() + "." + SmallArray4Names[0].Substring(0, 2) + SmallArray4Names[SmallArray4Names.Length - 1].Substring(0, 4) + AtCompany;
                        Console.WriteLine(completeMailAdress.ToLower());
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler: {0}", e.ToString());
            }//Email Adresse generieren aus der manuellen Eingabe

            using (FileStream Temp = new FileStream(PathToCSVFile, FileMode.Append, FileAccess.Write))                      //Instanziieren und Zugriff auf Methoden von FileStream zu Streamwriter
            using (StreamWriter CSV = new StreamWriter(Temp))                                                               // Zeilenweise in die Datei schreiben mit klein Buchstaben
            {
                CSV.WriteLine(EnteredName);
                CSV.WriteLine(completeMailAdress.ToLower());
                Console.WriteLine("Email Adresse wurde generiert und in {0} gespeichert", PathToCSVFile);
            }                                                                                                               // Zeilenweise in die Datei schreiben mit klein Buchstaben

            try//Inhalt CSV Datei ausgeben
            {
                using (StreamReader DeineDatei = File.OpenText(PathToCSVFile))
                {
                    string Inhalt3;
                    Console.WriteLine("Inhalt wird ausgegeben");
                    while ((Inhalt3 = DeineDatei.ReadLine()) != null)
                    {
                        Console.WriteLine(Inhalt3);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei der Durchfuehrung: {0}", e.ToString());
            }                                                                                                          //Inhalt CSV Datei ausgeben

            try
            {
                Console.WriteLine("Vordefinierter Email Versand wird gestartet.");
                var sender = new SmtpSender(() => new SmtpClient(host: "localhost")
                {
                    EnableSsl = false,
                    DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
                    PickupDirectoryLocation = Pfad

                });                                        // Testweise email versand mit speichern der Mail, fuer echten Versand smtp aendern und port, ggf 2 Factor Authentifizierung ändern
                Email.DefaultSender = sender;
                string htmlString = @"
                      
                    Some Text here" + Environment.NewLine +
                    "Best regards" + Environment.NewLine +
                    "John Doe";
                var email = await Email
                    .From(emailAddress: "someEmail@email.com")
                    .To(emailAddress: "someEmail@email.com", name: "John Doe")
                    .Subject(subject: "Betreffzeile: Test email zum speichern in Pfad")
                    //.Body = htmlString;
                    .Body(body: htmlString)



                    .SendAsync();
                Console.WriteLine("Email Versand wurde gespeichert in {0}", Pfad);
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler beim Email Versand: {0}", e.ToString());
            }                                                                                                         //Email Test Beispiel



            int looper = 0;
            while (looper == 0)//Ordner loeschen Abfrage mit while Schleife
            {
                Console.WriteLine("Moechten Sie die erstellten Ordner und Dateien loeschen? ja/nein");
                string Eingabe = Console.ReadLine();
                switch (Eingabe.ToLower())
                {
                    case "ja":
                        try
                        {
                            //Directory.CreateDirectory(Pfad);
                            Directory.Delete(Pfad, true);
                            bool directoryExists = Directory.Exists(Pfad);
                            bool subDirectoryExists = Directory.Exists(Pfad);

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Loeschen fehlgeschlagen. Fehler: {0}", e.Message);
                        }

                        Console.WriteLine("Ornder Geloescht");
                        looper = 1;
                        break;
                    case "nein":
                        Console.WriteLine("ok");
                        Console.WriteLine("Ordner wird zur Einsicht geöffnet");
                        Process.Start("explorer.exe", Pfad);

                        looper = 1;
                        break;
                    default:
                        Console.WriteLine("Falsche Eingabe.");
                        break;
                }
            }                                                                                          //Ordner loeschen Abfrage mit while Schleife
                                                                                                       //Thread.Sleep(5000);  //Waehrend der Entwicklung auf Kommentar setzen
                                                                                                       //Console.Clear();      //Waehrend der Entwicklung auf Kommentar setzen


            Console.WriteLine("\n\r\t\tPres any key to exit...");                                                           //Neue Zeile, Zeilenanfang, Tab Abtand, Tab Abstand - 
            Console.ReadKey();
        }


    }

}

