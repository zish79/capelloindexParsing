using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HtmlAgilityPack;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
                public    partial    class    Form1    :    Form    
                {    
                                public    Form1()    
                                {    
                                                InitializeComponent();    
                                                label10.Text    =   "League Name";  
                                                comboBox1.Items.Add("Premier League");  
                                                comboBox1.Items.Add("La Liga");  
                                                comboBox1.Items.Add("Bundesliga");    
                                                comboBox1.Items.Add("Ligue1");  
                                                comboBox1.Items.Add("Serie A");  
                                                comboBox1.Items.Add("Champions League");  
                                                comboBox1.Items.Add("Europa League");  
                                }

                                private void button1_Click_1(object sender, EventArgs e)
                                {    
                                                string    Drive    =    "";    
                                                string    file    =    "";    
                                                if    (!string.IsNullOrEmpty(textBox2.Text))    
                                                                Drive    =   textBox2.Text.ToString().Trim();  
                                                else    
                                                {    
                                                                MessageBox.Show("Enter    Drive   Name");  
                                                                return;    
                                                }    
                                                if    (!string.IsNullOrEmpty(textBox1.Text))    
                                                                file    +=   textBox1.Text.ToString().Trim();  
                                                else    
                                                {    
                                                                MessageBox.Show("Enter    File    Name");    
                                                                return;    
                                                }    
                                                //string    driveFile   =   Drive+":\";  
                                                //string   MyFile    =@Drive    +    ":\\"    +    file    +    "hh.xlsx";  
                                                string    MyFile    =    string.Format(@"{0}{1}{2}.xlsx",    Drive,":\\"  ,    file);  
                                                label14.Text    =    MyFile;  
                                                Microsoft.Office.Interop.Excel.Application    xl    =    null;    
 


                                                Microsoft.Office.Interop.Excel._Workbook    wb    =    null;    
                                                Microsoft.Office.Interop.Excel._Worksheet    sheet    =    null;    
    
                                                Microsoft.Office.Interop.Excel._Worksheet    sheet1    =    null;    
                                                Microsoft.Office.Interop.Excel._Worksheet    sheet2    =    null;    
                                                Microsoft.Office.Interop.Excel._Worksheet    sheet3    =    null;    
                                                Microsoft.Office.Interop.Excel._Worksheet    sheet4    =    null;    
    
                                                if    (File.Exists(MyFile))    {   File.Delete(MyFile);     }  
    
                                                xl    =    new    Microsoft.Office.Interop.Excel.Application();    
                                                xl.DisplayAlerts    =    false;    
                                                xl.Visible    =    false;    
    
                                                wb    =    (Microsoft.Office.Interop.Excel._Workbook)(xl.Workbooks.Add(Missing.Value));    
                                                wb.Sheets.Add(Missing.Value,    Missing.Value,   Missing.Value,   Missing.Value);    
                                                wb.Sheets.Add(Missing.Value,    Missing.Value,   Missing.Value,   Missing.Value);    
                                                wb.Sheets.Add(Missing.Value,    Missing.Value,   Missing.Value,   Missing.Value);    
                                                wb.Sheets.Add(Missing.Value,    Missing.Value,   Missing.Value,   Missing.Value);    
                                                wb.Sheets.Add(Missing.Value,    Missing.Value,   Missing.Value,   Missing.Value);    
    
                                                sheet    =   (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[1]);  
                                                sheet1    =   (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[2]);  
                                                sheet2    =   (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[3]);  
                                                sheet3    =   (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[4]);  
                                                sheet4    =   (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[5]);  
    
                                                sheet.Name    =   "Game   Data";  
                                                sheet1.Name    =   "Player   Data";  
                                                sheet2.Name    =   "Feuille    De    Match";    
                                                sheet3.Name    =   "Player    Stats";  
                                                sheet4.Name    =   "Player    Stats    1";    
    
                                                sheet.Cells[1]    =    "Id    Game";    
                                                sheet.Cells[2]    =    "Date";    
                                                sheet.Cells[3]    =    "Season";    
                                                sheet.Cells[4]    =    "Hour";    
                                                sheet.Cells[5]    =    "Competition";    
                                                sheet.Cells[6]    =   "Home   Team";  
                                                sheet.Cells[7]    =   "Home    Team    initial    League";    
                                                sheet.Cells[8]    =   "Away   Team";  
                                                sheet.Cells[9]    =    "Away    Team    initial    Leafue";    
                                                sheet.Cells[10]    =    "Home   Team   grade   on   the   game";    
                                                sheet.Cells[11]    =    "Away   Team   grade   on   the   game";    
                                                sheet.Cells[12]    =    "Home    Team    score";    
                                                sheet.Cells[13]    =    "Away    Team    Score";    
                                                sheet.Cells[14]    =    "URL   of   the   Data";    
                                                sheet.Cells[15]    =   "Date    of    Extraction";    
    
                                                sheet1.Cells[1]    =    "Id    Player   Data";    
                                                sheet1.Cells[2]    =    "Nom";    
                                                sheet1.Cells[3]    =    "Team";    
                                                sheet1.Cells[4]    =    "Position";    
                                                sheet1.Cells[5]    =    "Season";    
                                                sheet1.Cells[6]    =    "Birthday";    
                                                sheet1.Cells[7]    =    "Birthplace";    
                                                sheet1.Cells[8]    =    "Nationality";    
 


                                                sheet1.Cells[9]    =    "Weight";    
                                                sheet1.Cells[10]    =    "Height";    
                                                sheet1.Cells[11]    =    "Shirt";    
                                                sheet1.Cells[12]    =    "URL   of   the   Data";    
                                                sheet1.Cells[13]    =    "Date   of   Extraction";    
    
                                                sheet2.Cells[1]    =    "Id    Feuille   de   Match";    
                                                sheet2.Cells[2]    =    "Id    Match";    
                                                sheet2.Cells[3]    =    "Date";    
                                                sheet2.Cells[4]    =    "Season";    
                                                sheet2.Cells[5]    =    "Hour";    
                                                sheet2.Cells[6]    =    "Competition";    
                                                sheet2.Cells[7]    =   "Home   Team";  
                                                sheet2.Cells[8]    =    "Home    Team    initial    League";    
                                                sheet2.Cells[9]    =   "Away   Team";  
                                                sheet2.Cells[10]    =    "Away    Team    initial    Leafue";    
                                                sheet2.Cells[11]    =    "Player    name";    
                                                sheet2.Cells[12]    =    "Position";    
                                                sheet2.Cells[13]    =    "Team";    
                                                sheet2.Cells[14]    =   "Total    minutes    played";    
                                                sheet2.Cells[15]    =    "Grade   on   the   game";    
                                                sheet2.Cells[16]    =    "URL   of   the   Data";    
                                                sheet2.Cells[17]    =    "Date    of    Extraction";    
                                                DateTime    dt2    =    new    DateTime(2013,    05,    05,    00,    00,    0);  
                                                sheet3.Cells[1]    =    "Id    Player    stats";    
                                                sheet3.Cells[2]    =    "Id    Match";    
                                                sheet3.Cells[3]    =    "Nom";    
                                                sheet3.Cells[4]    =    "Team";    
                                                sheet3.Cells[5]    =    "Position";    
                                                sheet3.Cells[6]    =    "Team";    
                                                sheet3.Cells[7]    =    "Match";    
                                                sheet3.Cells[8]    =    "Compétition";    
                                                sheet3.Cells[9]    =    "CI    Value";    
                                                sheet3.Cells[10]    =    "Imp";    
                                                sheet3.Cells[11]    =    "Time";    
                                                sheet3.Cells[12]    =    "Event";    
                                                sheet3.Cells[13]    =   "Fields   Zone";  
                                                sheet3.Cells[14]    =    "URL   of   the   Data";    
                                                sheet3.Cells[15]    =    "Date    of    Extraction";    
    
                                                sheet4.Cells[1]    =    "Id    Player    stats";    
                                                sheet4.Cells[2]    =    "Id    Match";    
                                                sheet4.Cells[3]    =    "Nom";    
                                                sheet4.Cells[4]    =    "Team";    
                                                sheet4.Cells[5]    =    "Position";    
                                                sheet4.Cells[6]    =    "Team";    
                                                sheet4.Cells[7]    =    "Match";    
                                                sheet4.Cells[8]    =    "Compétition";    
                                                sheet4.Cells[9]    =    "CI    Value";    
                                                sheet4.Cells[10]    =    "Imp";    
                                                sheet4.Cells[11]    =    "Time";    
                                                sheet4.Cells[12]    =    "Event";    
                                                sheet4.Cells[13]    =   "Fields   Zone";  
                                                sheet4.Cells[14]    =    "URL   of   the   Data";    
                                                sheet4.Cells[15]    =    "Date    of    Extraction";    
    
                                     /*                                                                Work    Starts    from    Here                                                                                                                                                            
*/  
 


    
                                                 
                                                
                                                var    webGet    =    new   HtmlAgilityPack.HtmlWeb();    
    
    
                                                int    ioException    =    0;  
                                                int    webException    =    0;  
                                                int    Exception    =    0;  
    
    
                                                string[]    fr    =    new    string[7];
                                                fr[0] = "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=8&matchday=6";
//premier  Legue  
                                                fr[1]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=23&matchday=7";
//la    liga  
                                                fr[2]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=22&matchday=1";
//   Bundesliga  
                                                fr[3]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=24&matchday=1";
//Ligue1  
                                                fr[4]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=21&matchday=1";
//Seire    A  
                                                fr[5]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=5&matchday=1";
//Champions  Legue  
                                                fr[6]    =    "http://www.capelloindex.com/en/rankings?_dummy=0&detailtype=teams&league=6&matchday=1";
//Euoropa  Legue  
    
    
    
                                                //var    linksOnPage1;  
                                                int    Id_Match    =    1;  
                                                int    rows    =    2;  
                                                int    Player_Id    =    1;  
                                                int    rows_PlayerData    =    2;  
                                                int    id_player_stats    =    1;  
                                                int    row_player_stats    =    2;  
                                                int    row_player_stats1    =    2;  
                                                int    row_Feuille_Match    =    2;  
                                                int    id_Feullie    =    1;  
                                                string    Player_name    =    "",    Player_team    =    "",    Player_Position    =    "",    League    =    "";    
                                                string    Team1    =    "",    Team2    =    "",    Date    =    "",    Season    =    "",    Detail_Date_Data    =    "",    Time    =    "",    Score1    =    "",    Score2    =    "",    Grade1    =    "",    Grade2    =    "",    Player_URL    =    "";    
                                                int    current    =    0;  
                                                if    (comboBox1.SelectedIndex    ==   0)  
                                                                current    =    0;  
                                                if    (comboBox1.SelectedIndex    ==   1)  
                                                                current    =    1;  
                                                if    (comboBox1.SelectedIndex    ==   2)  
                                                                current    =    2;  
                                                if    (comboBox1.SelectedIndex    ==   3)  
                                                                current    =    3;  
 


                                                if    (comboBox1.SelectedIndex    ==   4)  
                                                                current    =    4;  
                                                if    (comboBox1.SelectedIndex    ==   5)  
                                                                current    =    5;  
                                                if    (comboBox1.SelectedIndex    ==   6)  
                                                                current    =    6;  
    
                                                try    
                                                {
                                                    ////webGet.PreRequest += request =>
                                                    ////{
                                                    ////    request.CookieContainer = new System.Net.CookieContainer();
                                                    ////    return true;
                                                    ////};
                                                               var    document    =   webGet.Load(fr[current]);   
                                                    if    (document.DocumentNode.SelectNodes("//ul[@class='pager paging']")    !=  null)  
                                                                {    
                                                                                var    hjhjh    =    from    lnks    in
                                                                                                         document.DocumentNode.SelectNodes("//ul[@class='pager paging']/li/a")  
                                                                                                                                where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                lnks.InnerText.Trim().Length    >   0  
                                                                                                                                select    new    
                                                                                                                                {    
                                                                                                                                                Url    =   lnks.Attributes["href"].Value,   
                                                                                                                                                Text    =    lnks.InnerText  
                                                                                                                                };    
                                                                                foreach    (var    itemJ    in    hjhjh)    
                                                                                {    
    
    
                                                                                                try    
                                                                                                {    
                                                                                                                if    (comboBox1.SelectedItem.ToString()    !=   null)  
                                                                                                                {    
                                                                                                                                string    LaLiga_URL    =    "http://www.capelloindex.com";    
                                                                                                                                //League    =    item.Text;  
                                                                                                                                //label10.Text    =    League;  
                                                                                                                                LaLiga_URL    +=    itemJ.Url;  
                                                                                                                                try    
                                                                                                                                {    
                                                                                                                                                document    =   webGet.Load(LaLiga_URL);  
                                                                                                                                }    
                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset   You    Internet");    document    =    webGet.Load(LaLiga_URL);    }  
                                                                                                                                catch    (IOException)    {    ioException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(LaLiga_URL);    }  
                                                                                                                                catch    (Exception)    {    Exception++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(LaLiga_URL);    }  
    
                                                                                                                                try    
                                                                                                                                {    
                                                                                                                                                var    linksOnPage1    =    from    lnks    in
                                                                                                                                                                                document.DocumentNode.SelectNodes("//ul[@class='playerList playerListTeam']/li/a")  
                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                lnks.Attributes["href"]    !=    
null  &&    
                                                                                                                                                                                                                                                lnks.InnerText.Trim().Length    
>    0  
 


                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                            Text    =    lnks.InnerText    
                                                                                                                                                                                                                            };    
                                                                                                                                                foreach    (var    club    in    linksOnPage1)    
                                                                                                                                                {    
                                                                                                                                                                try    
                                                                                                                                                                {    
                                                                                                                                                                                string    each_club    =    "http://www.capelloindex.com";                                                                            //    Select    Club    URL    from    the    List    of    Clubs    List  
    
                                                                                                                                                                                each_club    +=    club.Url;  
                                                                                                                                                                                try    
                                                                                                                                                                                {    
                                                                                                                                                                                                document    =   webGet.Load(each_club);  
                                                                                                                                                                                }    
    
                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(each_club);    }  
                                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(each_club);    }  
                                                                                                                                                                                catch    (Exception)    {    Exception++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(each_club);    }  
    
                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//table[@class='matchList']/tbody/tr/td/a")  
                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                                lnks.Attributes["href"]    
!=    null    &&    
                                                                                                                                                                                                                                                                
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                                            };    
    
    
    
    
    
    
                                                                                                                                                                                foreach    (var    match    in    linksOnPage1)                                                                                        
//    Match    Detail    Data  
                                                                                                                                                                                {    
    
                                                                                                                                                                                                
//***********************************************Game    Data    Sheet******************************************************//    
    
                                                                                                                                                                                                try    
 


                                                                                                                                                                                                {    
                                                                                                                                                                                                                label1.Text    =   Id_Match.ToString();  
                                                                                                                                                                                                                label4.Text    =    
webException.ToString();  
                                                                                                                                                                                                                label6.Text    =   ioException.ToString();  
                                                                                                                                                                                                                label8.Text    =   Exception.ToString();  
                                                                                                                                                                                                                sheet.Cells[rows,    1]    =    Id_Match;  
                                                                                                                                                                                                                //sheet.Cells[rows,    5]    =    League;  
                                                                                                                                                                                                                Season    =    "";    
                                                                                                                                                                                                                Date    =    "";    
                                                                                                                                                                                                                Time    =    "";    
                                                                                                                                                                                                                string    Match_clubs    =    "http://www.capelloindex.com";                                                                    //    Slect   Match    URL    of    club    Matches    List    
                                                                                                                                                                                                                Match_clubs    +=    match.Url;  
                                                                                                                                                                                                                int    matchID    =    
getPlayerID(Match_clubs);  
                                                                                                                                                                                                                sheet.Cells[rows,    14]    =    Match_clubs;  
                                                                                                                                                                                                                sheet.Cells[rows,    15]    =    DateTime.Now;    
                                                                                                                                                                                                                try    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                document    =    
webGet.Load(Match_clubs);  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet  Connection...Press    Ok    to    Continue   But   First   Reset   You   Internet");    document    =    webGet.Load(Match_clubs);     }  
                                                                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet     Connection...Press     Ok   to   Continue  But   First    Reset   You    Internet");    document    =    webGet.Load(Match_clubs);    }  
                                                                                                                                                                                                                catch    (Exception)    {    Exception++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(Match_clubs);    }  
    
                                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aLeagueDetail2']")  
                                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    
&&  
                                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                            Text    =    
lnks.InnerText  
                                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                          
//League  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                label10.Text    =    League    =    
sheet.Cells[rows,  5]   =   team.Text.ToString().Trim();  
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aTeamA']")  
 


                                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    
&&  
                                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                            Text    =    
lnks.InnerText  
                                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                          
//Home    Team  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                label11.Text    =    Team1    =    
sheet.Cells[rows,  6]   =   team.Text.ToString().Trim();  
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aTeamB']")  
                                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    
&&  
                                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                            Text    =    
lnks.InnerText  
                                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                          
//Away    Team  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                label13.Text    =    Team2    =    
sheet.Cells[rows,  8]   =   team.Text.ToString().Trim();  
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                var    Time_Detail    =    from    foo    in    
document.DocumentNode.SelectNodes("//span[@class='pubdate']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    itemB    in    Time_Detail)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Detail_Date_Data    =    
itemB.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                char[]    array    =    Detail_Date_Data.ToCharArray();    
                                                                                                                                                                                                                int    string_length    =   array.Length;  
                                                                                                                                                                                                                int    ii;    
                                                                                                                                                                                                                int    last    =    Detail_Date_Data.IndexOf(',');    
                                                                                                                                                                                                                for    (ii    =    last    +    1;    ii    <    
string_length;  ii++)  
 


                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Time    +=    array[ii];  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                sheet.Cells[rows,    4]    =    Time.ToString().Trim();    
                                                                                                                                                                                                                
                                                                                                                                                                                                    for  (ii=last - 5; ii < last;ii++)  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Season    +=    array[ii];  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                sheet.Cells[rows,    3]    =    Season.ToString().Trim();    
    
                                                                                                                                                                                                                for    (ii    =    0;    ii    <    last;    ii++)  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Date    +=    array[ii];  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                sheet.Cells[rows,    2]    =    Date.Trim();  
    
                                                                                                                                                                                                                int    score_C    =    1;  
                                                                                                                                                                                                                var    Score_Detail    =    from    foo    in    
document.DocumentNode.SelectNodes("//span[@class='result']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    itemC    in    Score_Detail)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                if    (score_C    ==    1)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                Score1    =    
itemC.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                if    (score_C    ==    2)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                Score2    =    
itemC.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                score_C++;    
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                sheet.Cells[rows,    12]    =    Score1;  
                                                                                                                                                                                                                sheet.Cells[rows,    13]    =    Score2;  
                                                                                                                                                                                                                var    Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//div[@class='teamCol   teamCol1']/div[@class='matchHeader']/span[@class='cindex']/span[@class='score']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    itemD    in    Grade_Detail)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Grade1    =    
itemD.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//div[@class='teamCol   teamCol2']/div[@class='matchHeader']/span[@class='cindex']/span[@class='score']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    itemE    in    Grade_Detail)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Grade2    =    
itemE.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                sheet.Cells[rows,    10]    =    Grade1;  
                                                                                                                                                                                                                sheet.Cells[rows,    11]    =    Grade2;  
    
 


    
    
                                                                                                                                                                                                                
//**************************************************Game    Data   Sheet    Done**************************************//  
    
                                                                                                                                                                                                                
//**************************************************Feuille    De   Matach  Data  Sheet  
//**************************************//  
                                                                                                                                                                                                                var    Player_List_Count    =    from    lnks    in document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a") where lnks.Name    ==    "a"    &&   lnks.Attributes["href"]    !=    null    &&  lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                                                                                select    new    
                                                                                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                                                                                Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                                                                Text    =    
lnks.InnerText,  
                                                                                                                                                                                                                                                                                                                                id    =    
lnks.Attributes["id"].Value  
                                                                                                                                                                                                                                                                                                                };    
    
                                                                                                                                                                                                                foreach    (var    itemG    in    Player_List_Count)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
1]    =    id_Feullie;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
2]    =    Id_Match;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
3]    =    Date;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
4]    =    Season;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
5]    =    Time;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
6]    =    League;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
7]    =    Team1;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
9]    =    Team2;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
16]    =    Match_clubs;  
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
17]    =    DateTime.Now;    
                                                                                                                                                                                                                                if    (itemG.id.Contains("matchPlayerListA"))    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    13]   =   Team1;  
 


                                                                                                                                                                                                                                                var    Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='name']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    11]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='role']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    12]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='score']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    15]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='label']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    14]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                else    if    (itemG.id.Contains("matchPlayerListB"))    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,  13]    =    Team2;  
                                                                                                                                                                                                                                                var    Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='name']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    Player_Detail_S)    
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    11]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='role']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    Player_Detail_S)    
                                                                                                                                                                                                                                                {    
 


                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    12]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='score']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    15]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='label']")    select    lnks;    
                                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    14]   =   itemH.InnerText;  
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                id_Feullie++;    
                                                                                                                                                                                                                                row_Feuille_Match++;    
                                                                                                                                                                                                                }    
    
    
    
    
    
                                                                                                                                                                                                                
//**************************************************Feuille    De   Matach  Data  Sheet    Done**************************************//  
    
    
    
                                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a")  
                                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    
&&  
                                                                                                                                                                                                                                                                                                
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                            Text    =    
lnks.InnerText  
                                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                                foreach    (var    itemF    in    linksOnPage1)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                label1.Text    =    Id_Match.ToString();    
                                                                                                                                                                                                                                label4.Text    =    
webException.ToString();  
 


                                                                                                                                                                                                                                label6.Text    =    
ioException.ToString();  
                                                                                                                                                                                                                                label8.Text    =    Exception.ToString();    
                                                                                                                                                                                                                                
//*********************************************************Player    Data    Sheet*********************************//  
                                                                                                                                                                                                                                Player_URL    =    "http://www.capelloindex.com";    
                                                                                                                                                                                                                                Player_URL    +=    itemF.Url;  
                                                                                                                                                                                                                                int    playerID    =    
getPlayerID(Player_URL);  
    
                                                                                                                                                                                                                                try    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                document    =    
webGet.Load(Player_URL);  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue   But   First   Reset   You   Internet");    document    =    webGet.Load(Player_URL);     }  
                                                                                                                                                                                                                                catch    (IOException)    {    
ioException++;  MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue   But   First   Reset   You   Internet");    document    =    webGet.Load(Player_URL);     }  
                                                                                                                                                                                                                                catch    (Exception)    {    Exception++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(Player_URL);    }  
    
    
                                                                                                                                                                                                                                var    Player_Data    =    Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//h1[@id='ctl00_cph_content_playerInfo1_h1PlayerName'] ")    select    foo;    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    1]   
=    Player_Id;  
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    12]   
=   Player_URL;  
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    13]   
=    DateTime.Now;    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    5]   
=    Season;  
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                Player_name    =    
sheet1.Cells[rows_PlayerData,    2]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    from    lnks    in    document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_playerInfo1_aTeam']/span[@c lass='team']")    select    lnks;    
    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                Player_team    =    
sheet1.Cells[rows_PlayerData,    3]    =    pla.InnerText.ToString().Trim();    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@class='role']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
 


                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                Player_Position    =    
sheet1.Cells[rows_PlayerData,    4]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@itemprop='birthDate']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
6]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@class='birthplace']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
7]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@itemprop='nationality']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
8]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@class='weight']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
9]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =   Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@class='height']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
10]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Data    =  Grade_Detail    =    from    
foo    in    document.DocumentNode.SelectNodes("//span[@class='shirt']")    select    foo;    
                                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    
11]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                                Player_Id++;    
                                                                                                                                                                                                                                rows_PlayerData++;    
    
                                                                                                                                                                                                                                
//*********************************************************Player    Data   Sheet    Done*********************************//  
    
    
                                                                                                                                                                                                                                
//*********************************************************Player    Stats   Data    Sheet*********************************//  
 


                                                                                                                                                                                                                                string    Json_Player_Stats_URL    =    "http://www.capelloindex.com/en/svc/fci.ashx?playerid="    +    playerID    +    "&matchid="    +    matchID;    
                                                                                                                                                                                                                                string    reply    =    "";    
                                                                                                                                                                                                                                WebClient    client    =    new    WebClient();    
                                                                                                                                                                                                                                try    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                reply    =    "{\"Size\""    +    ":"    +    
client.DownloadString(Json_Player_Stats_URL)  +  "}";    
                                                                                                                                                                                                                                                JObject    o    =    JObject.Parse(reply);    
                                                                                                                                                                                                                                                //string    name    = (string)o["name"];  
                                                                                                                                                                                                                                                JArray    sizes    =    (JArray)o["Size"];    
                                                                                                                                                                                                                                                //string    smallest    = sizes[0]["name"].ToString();  
                                                                                                                                                                                                                                                if    (row_player_stats <=1048576 - 4)    
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                for    (int    i    =    0;    i    <    
sizes.Count;  i++)  
                                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    1]   =   id_player_stats;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    2]   =   Id_Match;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    3]   =   Player_name;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    4]   =   Player_team;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    5]   =   Player_Position;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    6]   =   Player_team;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    7]    =    Team1    +    "   vs.   "  +    Team2;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    8]   =   League;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    9]   =   sizes[i]["isPositive"].ToString();  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    10]   =   sizes[i]["importance"].ToString();  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    11]   =   sizes[i]["time"].ToString();  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    12]   =   sizes[i]["name"].ToString();  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    13]   =   sizes[i]["zone"].ToString();  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    14]   =   Player_URL;  
                                                                                                                                                                                                                                                                                
sheet3.Cells[row_player_stats,    15]    =    DateTime.Now;    
                                                                                                                                                                                                                                                                                id_player_stats++;    
                                                                                                                                                                                                                                                                                row_player_stats++;    
                                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                }    
 


                                                                                                                                                                                                                                                else    
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                for    (int    i    =    0;    i    <    
sizes.Count;  i++)  
                                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    1]   =   id_player_stats;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    2]   =   Id_Match;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    3]    =    Player_name;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    4]   =   Player_team;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    5]   =   Player_Position;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    6]   =   Player_team;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    7]    =    Team1    +    "   vs.   "  +    Team2;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    8]   =   League;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    9]   =   sizes[i]["isPositive"].ToString();  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    10]   =   sizes[i]["importance"].ToString();  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    11]   =   sizes[i]["time"].ToString();  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    12]   =   sizes[i]["name"].ToString();  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    13]   =   sizes[i]["zone"].ToString();  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    14]   =   Player_URL;  
                                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    15]  =   DateTime.Now;  
                                                                                                                                                                                                                                                                                id_player_stats++;    
                                                                                                                                                                                                                                                                                row_player_stats1++;    
                                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs   on   Internet   Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                                                catch    (IOException)    {    
ioException++;  MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First  Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                                                catch    (Exception)    {    Exception++;    
continue;    }  
    
                                                                                                                                                                                                                                
//*********************************************************Player    Stats    Data    Sheet    Done*********************************//    
    
    
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                rows++;    
                                                                                                                                                                                                                Id_Match++;    
 


                                                                                                                                                                                                                xl.Visible    =    false;    
                                                                                                                                                                                                                xl.UserControl    =    false;    
                                                                                                                                                                                                                wb.SaveAs(MyFile,    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,    Missing.Value,    
                Missing.Value,    false,    false,    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,    
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution,    true,  
                Missing.Value,    Missing.Value,   Missing.Value);    
                                                                                                                                                                                                }    
                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                catch    (Exception    ex)    {    Exception++;    MessageBox.Show(ex.Message);    continue;    }  
                                                                                                                                                                                }    
                                                                                                                                                                }    
                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                catch    (IOException)    {    ioException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue   But   First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                catch    (Exception)    {    Exception++;    continue;    }  
    
    
                                                                                                                                                }    
                                                                                                                                }    
                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                catch    (Exception)    {    Exception++;    continue;    }  
    
                                                                                                                }    
                                                                                                }    
                                                                                                catch    (WebException)    {  webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                catch    (IOException)    {    ioException++;    MessageBox.Show("Error Occurs   on   Internet   Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    
continue;    }  
                                                                                                catch    (Exception)    {    Exception++;    continue;    }  
                                                                                }    
                                                                }    
                                                                else    
                                                                {    
                                                                                try    
                                                                                {    
                                                                                                if    (comboBox1.SelectedItem.ToString()    !=   null)  
                                                                                                {    
    
                                                                                                                //League    =    item.Text;  
                                                                                                                //label10.Text    =    League;  
 


                                                                                                                try    
                                                                                                                {    
                                                                                                                                var    linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//ul[@class='playerList    playerListTeam']/li/a")  
                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                lnks.Attributes["href"]    !=    null    
&&  
                                                                                                                                                                                                                                lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                            {    
                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                            };    
                                                                                                                                foreach    (var    club    in    linksOnPage1)    
                                                                                                                                {    
                                                                                                                                                try    
                                                                                                                                                {    
                                                                                                                                                                string    each_club    =    "http://www.capelloindex.com";                                                
//    Select    Club    URL    from    the    List    of    Clubs    List  
    
                                                                                                                                                                each_club    +=    club.Url;  
                                                                                                                                                                try    
                                                                                                                                                                {    
                                                                                                                                                                                document    =   webGet.Load(each_club);  
                                                                                                                                                                }    
                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to  Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(each_club);    }  
                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset  You   Internet");    document   =   webGet.Load(each_club);  }  
                                                                                                                                                                catch    (Exception)    {    Exception++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(each_club);    }  
    
    
    
                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//table[@class='matchList']/tbody/tr/td/a")  
                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                lnks.Attributes["href"]    !=    
null  &&    
                                                                                                                                                                                                                                                lnks.InnerText.Trim().Length    
>    0  
                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                            };    
    
    
    
    
    
    
 


                                                                                                                                                                foreach    (var    match    in    linksOnPage1)                                                                                                        
//    Match    Detail    Data  
                                                                                                                                                                {    
    
                                                                                                                                                                                
//***********************************************Game    Data    Sheet******************************************************//    
    
                                                                                                                                                                                try    
                                                                                                                                                                                {    
                                                                                                                                                                                                label1.Text    =   Id_Match.ToString();  
                                                                                                                                                                                                label4.Text    =   webException.ToString();  
                                                                                                                                                                                                label6.Text    =   ioException.ToString();  
                                                                                                                                                                                                label8.Text    =   Exception.ToString();  
                                                                                                                                                                                                sheet.Cells[rows,    1]    =    Id_Match;  
                                                                                                                                                                                                //sheet.Cells[rows,    5]    =    League;  
                                                                                                                                                                                                Season    =    "";    
                                                                                                                                                                                                Date    =    "";    
                                                                                                                                                                                                Time    =    "";    
                                                                                                                                                                                                string    Match_clubs    =    "http://www.capelloindex.com";                                                                    //    Slect    Match    URL    of    club    Matches    List    
                                                                                                                                                                                                Match_clubs    +=    match.Url;  
                                                                                                                                                                                                int    matchID    =   getPlayerID(Match_clubs);  
                                                                                                                                                                                                sheet.Cells[rows,    14]    =    Match_clubs;  
                                                                                                                                                                                                sheet.Cells[rows,    15]    =    DateTime.Now;    
                                                                                                                                                                                                try    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                document    =   webGet.Load(Match_clubs);  
                                                                                                                                                                                                }    
                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset   You    Internet");    document    =    webGet.Load(Match_clubs);    }  
                                                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet     Connection...Press     Ok   to   Continue  But   First    Reset    You    Internet");    document    =    webGet.Load(Match_clubs);     }  
                                                                                                                                                                                                catch    (Exception)    {    Exception++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You   Internet");   document    =    webGet.Load(Match_clubs);    }  
    
    
                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aLeagueDetail2']")  
                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                                           
//League  
                                                                                                                                                                                                {    
                                                                                                                                                                                                                label10.Text    =    League    =    
sheet.Cells[rows,  5]   =   team.Text.ToString().Trim();  
 


                                                                                                                                                                                                }    
    
                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aTeamA']")  
                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                                           
//Home    Team  
                                                                                                                                                                                                {    
                                                                                                                                                                                                                label11.Text    =    Team1    =    
sheet.Cells[rows,  6]   =   team.Text.ToString().Trim();  
                                                                                                                                                                                                }    
    
                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_aTeamB']")  
                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                                                            
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                foreach    (var    team    in    linksOnPage1)                                                                           
//Away    Team  
                                                                                                                                                                                                {    
                                                                                                                                                                                                                label13.Text    =    Team2    =    
sheet.Cells[rows,  8]   =   team.Text.ToString().Trim();  
                                                                                                                                                                                                }    
    
                                                                                                                                                                                                var    Time_Detail    =    from    foo    in    
document.DocumentNode.SelectNodes("//span[@class='pubdate']")    select    foo;    
                                                                                                                                                                                                foreach    (var    itemB    in    Time_Detail)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                Detail_Date_Data    =    
itemB.InnerText.ToString().Trim();  
                                                                                                                                                                                                }    
    
                                                                                                                                                                                                char[]    array    =    Detail_Date_Data.ToCharArray();    
                                                                                                                                                                                                int    string_length    =   array.Length;  
                                                                                                                                                                                                int    ii;    
                                                                                                                                                                                                int    last    =    Detail_Date_Data.IndexOf(',');    
                                                                                                                                                                                                for    (ii    =    last    +    1;    ii    <    string_length;    
ii++)  
                                                                                                                                                                                                {    
 


                                                                                                                                                                                                                Time    +=    array[ii];  
                                                                                                                                                                                                }    
                                                                                                                                                                                                sheet.Cells[rows,    4]    =    Time.ToString().Trim();    
                                                                                                                                                                                                for    (ii = last -5; ii  <  last;    ii++)  
                                                                                                                                                                                                {    
                                                                                                                                                                                                                Season    +=    array[ii];  
                                                                                                                                                                                                }    
                                                                                                                                                                                                sheet.Cells[rows,    3]    =    Season.ToString().Trim();    
    
                                                                                                                                                                                                for    (ii    =    0;    ii    <    last;    ii++)  
                                                                                                                                                                                                {    
                                                                                                                                                                                                                Date    +=    array[ii];  
                                                                                                                                                                                                }    
                                                                                                                                                                                                sheet.Cells[rows,    2]    =    Date.Trim();  
    
                                                                                                                                                                                                int    score_C    =    1;  
                                                                                                                                                                                                var    Score_Detail    =    from    foo    in    
document.DocumentNode.SelectNodes("//span[@class='result']")    select    foo;    
                                                                                                                                                                                                foreach    (var    itemC    in    Score_Detail)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                if    (score_C    ==    1)  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Score1    =    
itemC.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                if    (score_C    ==    2)  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Score2    =    
itemC.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                score_C++;    
                                                                                                                                                                                                }    
                                                                                                                                                                                                sheet.Cells[rows,    12]    =    Score1;  
                                                                                                                                                                                                sheet.Cells[rows,    13]    =    Score2;  
                                                                                                                                                                                                var    Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//div[@class='teamCol   teamCol1']/div[@class='matchHeader']/span[@class='cindex']/span[@class='score']")    select    foo;    
                                                                                                                                                                                                foreach    (var    itemD    in    Grade_Detail)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                Grade1    =    
itemD.InnerText.ToString().Trim();  
                                                                                                                                                                                                }    
                                                                                                                                                                                                Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//div[@class='teamCol   teamCol2']/div[@class='matchHeader']/span[@class='cindex']/span[@class='score']")    select    foo;    
                                                                                                                                                                                                foreach    (var    itemE    in    Grade_Detail)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                Grade2    =    
itemE.InnerText.ToString().Trim();  
                                                                                                                                                                                                }    
                                                                                                                                                                                                sheet.Cells[rows,    10]    =    Grade1;  
                                                                                                                                                                                                sheet.Cells[rows,    11]    =    Grade2;  
    
    
 


    
                                                                                                                                                                                                
//**************************************************Game    Data    Sheet    Done**************************************//    
    
                                                                                                                                                                                                
//**************************************************Feuille    De   Matach  Data  Sheet  
//**************************************//  
    
                                                                                                                                                                                                var    Player_List_Count    =    from    lnks    in    
document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a")  
                                                                                                                                                                                                                                                                                                where    lnks.Name    
==    "a"    &&    
                                                                                                                                                                                                                                                lnks.Attributes["href"]    !=    
null  &&    
                                                                                                                                                                                                                                                lnks.InnerText.Trim().Length    
>    0  
                                                                                                                                                                                                                                                                                                select    new    
                                                                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                                                                                Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                                                                Text    =    
lnks.InnerText,  
                                                                                                                                                                                                                                                                                                                id    =    
lnks.Attributes["id"].Value  
                                                                                                                                                                                                                                                                                                };    
    
                                                                                                                                                                                                foreach    (var    itemG    in    Player_List_Count)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    1]    =    
id_Feullie;  
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    2]    =    Id_Match;    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    3]    =    Date;    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    4]    =    
Season;  
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    5]    =    Time;    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    6]    =    League;    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    7]    =    Team1;    
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    9]    =    
Team2;  
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    16]  =    Match_clubs;  
                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    17]  =    DateTime.Now;  
                                                                                                                                                                                                                if    
(itemG.id.Contains("matchPlayerListA"))  
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
13]    =    Team1;  
                                                                                                                                                                                                                                var    Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='name']")    select    lnks;    
 


                                                                                                                                                                                                                                foreach    (var    itemH    in    Player_Detail_S)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    11]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='role']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    12]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='score']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    15]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='label']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    14]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                else    if    (itemG.id.Contains("matchPlayerListB"))    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet2.Cells[row_Feuille_Match,    
13]    =    Team2;  
                                                                                                                                                                                                                                var    Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='name']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    Player_Detail_S)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    11]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='info']/span[@class='role']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    Player_Detail_S)    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    12]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
 


                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='score']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    15]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                Player_Detail_S    =    from    lnks    in    document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a[@id='"    +    itemG.id    +    "']/span[@class='cindex']/span[@class='label']")    select    lnks;    
                                                                                                                                                                                                                                foreach    (var    itemH    in    
Player_Detail_S)  
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                
sheet2.Cells[row_Feuille_Match,    14]   =   itemH.InnerText;  
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                id_Feullie++;    
                                                                                                                                                                                                                row_Feuille_Match++;    
                                                                                                                                                                                                }    
    
    
    
    
    
                                                                                                                                                                                                
//**************************************************Feuille    De    Matach    Data    Sheet    Done**************************************//    
    
    
    
                                                                                                                                                                                                linksOnPage1    =    from    lnks    in    
document.DocumentNode.SelectNodes("//ul[@class='playerList']/li/a")  
                                                                                                                                                                                                                                                            where    lnks.Name    ==    "a"    &&    
                                                                                                                                                                                                                                                                                
lnks.Attributes["href"]    !=    null    &&    
                                                                                                                                                                                                                                                                                
lnks.InnerText.Trim().Length    >   0  
                                                                                                                                                                                                                                                            select    new    
                                                                                                                                                                                                                                                            {    
                                                                                                                                                                                                                                                                            Url    =    
lnks.Attributes["href"].Value,  
                                                                                                                                                                                                                                                                            Text    =    lnks.InnerText  
                                                                                                                                                                                                                                                            };    
                                                                                                                                                                                                foreach    (var    itemF    in    linksOnPage1)    
                                                                                                                                                                                                {    
                                                                                                                                                                                                                label1.Text    =   Id_Match.ToString();  
                                                                                                                                                                                                                label4.Text    =    
webException.ToString();  
                                                                                                                                                                                                                label6.Text    =   ioException.ToString();  
                                                                                                                                                                                                                label8.Text    =   Exception.ToString();  
                                                                                                                                                                                                                
//*********************************************************Player    Data    Sheet*********************************//  
                                                                                                                                                                                                                Player_URL    =    "http://www.capelloindex.com";    
 


                                                                                                                                                                                                                Player_URL    +=   itemF.Url;  
                                                                                                                                                                                                                    
                                                                                                                                                                                                                int    playerID    =    
getPlayerID(Player_URL);  
    
                                                                                                                                                                                                                try    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                document    =    
webGet.Load(Player_URL);  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue   But   First   Reset   You   Internet");    document    =    webGet.Load(Player_URL);     }  
                                                                                                                                                                                                                catch    (IOException)    {    ioException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    document    =    webGet.Load(Player_URL);    }  
                                                                                                                                                                                                                catch    (Exception)    {    Exception++;    MessageBox.Show("Error    Occurs    on    Internet     Connection...Press     Ok   to   Continue  But   First    Reset    You    Internet");    document    =    webGet.Load(Player_URL);     }  
                                                                                                                                                                                                                var    Player_Data    =    Grade_Detail    =    from    foo    in    document.DocumentNode.SelectNodes("//h1[@id='ctl00_cph_content_playerInfo1_h1PlayerName'] ")    select    foo;    
                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    1]    =    
Player_Id;  
                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    12]    =    Player_URL;    
                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    13]    =    DateTime.Now;    
                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    5]    =    
Season;  
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Player_name    =    
sheet1.Cells[rows_PlayerData,    2]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    from    lnks    in    document.DocumentNode.SelectNodes("//a[@id='ctl00_cph_content_playerInfo1_aTeam']/span[@c lass='team']")    select    lnks;    
    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Player_team    =    
sheet1.Cells[rows_PlayerData,    3]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@class='role']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                Player_Position    =    
sheet1.Cells[rows_PlayerData,    4]   =   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@itemprop='birthDate']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    6]   
=   pla.InnerText.ToString().Trim();  
 


                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@class='birthplace']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    7]    
=   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@itemprop='nationality']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    8]    
=   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@class='weight']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    9]    
=   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@class='height']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    10]    
=   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                Player_Data    =    Grade_Detail    =    from    foo    
in  document.DocumentNode.SelectNodes("//span[@class='shirt']")    select    foo;    
                                                                                                                                                                                                                foreach    (var    pla    in    Player_Data)    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                sheet1.Cells[rows_PlayerData,    11]    
=   pla.InnerText.ToString().Trim();  
                                                                                                                                                                                                                }    
    
                                                                                                                                                                                                                Player_Id++;    
                                                                                                                                                                                                                rows_PlayerData++;    
    
                                                                                                                                                                                                                
//*********************************************************Player    Data   Sheet    Done*********************************//  
    
    
                                                                                                                                                                                                                
//*********************************************************Player    Stats   Data    Sheet*********************************//  
                                                                                                                                                                                                                string    Json_Player_Stats_URL    =    "http://www.capelloindex.com/en/svc/fci.ashx?playerid="    +    playerID    +    "&matchid="    +    matchID;  
                                                                                                                                                                                                                string    reply    =    "";    
                                                                                                                                                                                                                WebClient    client    =    new    WebClient();    
                                                                                                                                                                                                                try    
                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                reply    =    "{\"Size\""    +    ":"    +    
client.DownloadString(Json_Player_Stats_URL)    +   "}";  
                                                                                                                                                                                                                                JObject    o    =    JObject.Parse(reply);    
 


                                                                                                                                                                                                                                //string    name    =    (string)o["name"];    
                                                                                                                                                                                                                                JArray    sizes    =    (JArray)o["Size"];    
                                                                                                                                                                                                                                //string    smallest    =    
sizes[0]["name"].ToString();  
                                                                                                                                                                                                                                if    (row_player_stats <= 1048572  )  //1048576-4 brijesh
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                for    (int i =0;i<sizes.Count;  i++)  
                                                                                                                                                                                                                                               {    
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,  1]    =   id_player_stats;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    2]   =   Id_Match;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    3]   =   Player_name;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    4]   =   Player_team;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    5]   =   Player_Position;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    6]   =   Player_team;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    7]    =    Team1    +    "   vs.   "  +    Team2;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    8]   =   League;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    9]   =   sizes[i]["isPositive"].ToString();  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    10]  =   sizes[i]["importance"].ToString();  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    11]   =   sizes[i]["time"].ToString();  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    12]   =   sizes[i]["name"].ToString();  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    13]   =   sizes[i]["zone"].ToString();  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    14]   =   Player_URL;  
                                                                                                                                                                                                                                                  sheet3.Cells[row_player_stats,    15]    =    DateTime.Now;    
                                                                                                                                                                                                                                                  id_player_stats++;    
                                                                                                                                                                                                                                                  row_player_stats++;    
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                else    
                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                for    (int i = 0; i <sizes.Count;  i++)  
                                                                                                                                                                                                                                                {    
                                                                                                                                                                                                                                                     sheet4.Cells[row_player_stats1,    1]   =   id_player_stats;  
                                                                                                                                                                                                                                                     sheet4.Cells[row_player_stats1,    2]   =   Id_Match;  
                                                                                                                                                                                                                                                                
                                                sheet4.Cells[row_player_stats1,    3]   =   Player_name;  
                                                                                                                                                                                                                                                                
                                                sheet4.Cells[row_player_stats1,    4]   =   Player_team;  
 


                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    5]   =   Player_Position;  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    6]   =   Player_team;  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    7]    =    Team1    +    "   vs.   "  +    Team2;  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    8]   =   League;  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,  9]   =   sizes[i]["isPositive"].ToString();  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    10]   =   sizes[i]["importance"].ToString();  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    11]   =   sizes[i]["time"].ToString();  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    12]   =   sizes[i]["name"].ToString();  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    13]   =   sizes[i]["zone"].ToString();  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    14]   =   Player_URL;  
                                                                                                                                                                                                                                                                
sheet4.Cells[row_player_stats1,    15]  =   DateTime.Now;  
                                                                                                                                                                                                                                                                id_player_stats++;    
                                                                                                                                                                                                                                                                row_player_stats1++;    
                                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                                }    
                                                                                                                                                                                                                }    
                                                                                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                                                catch    (Exception)    {    Exception++;    
continue;    }  
    
                                                                                                                                                                                                                
//*********************************************************Player    Stats   Data   Sheet    Done*********************************//  
    
    
                                                                                                                                                                                                }    
    
                                                                                                                                                                                                rows++;    
                                                                                                                                                                                                Id_Match++;    
                                                                                                                                                                                                wb.SaveAs(MyFile,    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,   Missing.Value,Missing.Value,    false,    false,    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,    
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution,    true,Missing.Value,    Missing.Value,   Missing.Value);    
                                                                                                                                                                                }    
                                                                                                                                                                                catch    (WebException)    
                                                                                                                                                                                {    
                                                                                                                                                                                                webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    
                                                                                                                                                                                                continue;    
                                                                                                                                                                                }    
 


                                                                                                                                                                                catch    (IOException)    {    ioException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue   But   First    Reset    You    Internet");    continue;    }  
                                                                                                                                                                                catch    (Exception    ex)    {    Exception++;    
MessageBox.Show(ex.InnerException.ToString());  MessageBox.Show(ex.Message);  continue;    }  
                                                                                                                                                                }    
                                                                                                                                                }    
                                                                                                                                                catch    (WebException)    {    webException++;    MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                catch    (IOException)    {    ioException++;    
MessageBox.Show("Error    Occurs    on    Internet    Connection...Press    Ok    to    Continue    But    First    Reset    You    Internet");    continue;    }  
                                                                                                                                                catch    (Exception)    {    Exception++;    continue;    }  
    
                                                                                                                                }    
                                                                                                                }    
                                                                                                                catch    (WebException)    {    webException++;    }  
                                                                                                                catch    (IOException)    {    ioException++;    }  
                                                                                                                catch    (Exception)    {    Exception++;    }  
    
                                                                                                }    
                                                                                }    
                                                                                catch    (WebException)    
                                                                                {    
                                                                                                MessageBox.Show("Error    Occurs    on    Internet    Connection");    
webException++;  
                                                                                }    
                                                                                catch    (IOException)    
                                                                                {    
                                                                                                MessageBox.Show("Error    Occurs    on    Internet    Connection");    
ioException++;  
                                                                                }    
                                                                                catch    (Exception)    
                                                                                {    
                                                                                                Exception++;    
                                                                                }    
                                                                }    
                                                }    
                                                catch    (WebException)    
                                                {    
                                                                MessageBox.Show("Web    site    is    in    Error    condition,    So    try    Later");    
                                                }    
                                                catch    (IOException)    
                                                {    
                                                                MessageBox.Show("Web    site    is    in    Error    condition,    So    try    Later");    
                                                }    
    
    
                                                //********************************************************Premier    Ligue*********************************************************//  
    
    
    
                                                xl.Visible    =    false;    
                                                xl.UserControl    =    false;    
    
                                                //        Save    the    file    to    disk  
 


    
                                                //wb.SaveAs(MyFile,    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,    
                                                //                                        null,    null,    false,    false,    
//Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,  
                                                //                                        false,    false,    null,    null,    null);  
                                                wb.SaveAs(MyFile,    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,   Missing.Value,    
                Missing.Value,    false,    false,    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,    
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution,    true,  
                Missing.Value,    Missing.Value,   Missing.Value);    
                                                xl.Visible    =    false;    
                                                xl.UserControl    =    false;    
                                                //    Close    the    document    and    avoid    user    prompts    to    save    if    our    method    failed.  
                                                wb.Close(true,    null,    null);    
                                                xl.Workbooks.Close();    
                                                return;    
                                }    
                                public    int    getPlayerID(string    Player_URL)    
                                {    
                                                //      www.capelloindex.com/en/player/ben-­‐foster/9089    
                                                int    last    =   Player_URL.LastIndexOf('/');  
                                                char[]    array    =    Player_URL.ToArray();  
                                                string    aa    =    "";    
                                                for    (int    i    =   last  +   1;    i    <    array.Length;      i++)  
                                                {    
                                                                aa    +=    array[i];  
                                                }    
                                                return    Convert.ToInt32(aa);    
                                }

                    private void label16_Click(object sender, EventArgs e)
                    {
                    
                    }
                }    

}
