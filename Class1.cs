
using App = HostMgd.ApplicationServices;
using Db = Teigha.DatabaseServices;
using Ed = HostMgd.EditorInput;
using Rtm = Teigha.Runtime;
using System.Diagnostics;

using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;


[assembly: Rtm.CommandClass(typeof(Tools.CadCommand))]

namespace Tools
    {
        /// <summary> 
        /// Комманды
        /// </summary>
        class CadCommand : Rtm.IExtensionApplication
        {

            #region INIT
            public void Initialize()
            {

                App.DocumentCollection dm = App.Application.DocumentManager;
                Ed.Editor ed = dm.MdiActiveDocument.Editor;
                string sCom =
                    "trypdf" + "\tСоздание PDF из nanoCAD";
                ed.WriteMessage(sCom);

            }

            public void Terminate()
            {
                // выход из модуля
            }

            #endregion

            #region Command

            /// <summary>
            /// Проба pdfSharp
            /// </summary>
            [Rtm.CommandMethod("trypdf")]
            public void trypdf()
            {
                Db.Database db = Db.HostApplicationServices.WorkingDatabase;
                App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
                Ed.Editor ed = doc.Editor;
            try
                {
                genPDF.createPDF();
                }
                catch (Exception ex)
                {
                MessageBox.Show("Ошибка: " + ex.Message);
                }


        }

        #endregion

        class genPDF
        {
            public static void createPDF()
            {
                // Путь для сохранения
                string pathToSave = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filename = "Новый PDF-документ.pdf";
                string fullFilePath = $"{pathToSave}\\{filename}";

                // Создаем новый PDF документ
                PdfDocument document = new PdfDocument();
                document.Info.Title = "PDF файл созданный с помощью PDFsharp"; // Установка заголовка документа

                // Добавляем пустую страницу
                PdfPage page = document.AddPage();

                // Создаем графику для рисования на странице
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);

                // Добавляем текст на страницу (по желанию)
                gfx.DrawString("Я родился! Я существую!", font, XBrushes.Black,
                    new XRect(0, 0, page.Width, page.Height),
                    XStringFormats.Center);

                // Сохраняем документ в файл
                document.Save(fullFilePath);

                // Открытие созданного файла после его сохранения (по желанию)
                Process.Start(new ProcessStartInfo()
                {
                    FileName = fullFilePath,
                    UseShellExecute = true
                });

                MessageBox.Show($"PDF файл '{filename}' успешно создан.");
            }
        }


    }
    }


