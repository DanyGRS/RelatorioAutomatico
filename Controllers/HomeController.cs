using Microsoft.Office.Interop.Excel;
using System.Web.Mvc;

namespace teste.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult ExecutaMacro()
        {
            string caminho = "C:\\Users\\PC\\source\\repos\\teste\\teste\\Models\\certi.xlsm";
            Application xlApp = new Application();
         
            if (xlApp == null)
                {
                    ViewBag.Mensagem = "Erro ao executar a macro: aplicativo Excel não encontrado.";
                    return View("ChamarMacro");
                }

                Workbook xlWorkbook = xlApp.Workbooks.Open(caminho, ReadOnly: false);

                try
                {
                    xlApp.Visible = false;
                    xlApp.Run("GerarCertificado");
                }
                catch (System.Exception)
                {
                    ViewBag.Mensagem = "Erro ao executar a macro.";
                    return View("ChamarMacro");
                }

                xlWorkbook.Close(false);
                xlApp.Application.Quit();
                xlApp.Quit();
            

            ViewBag.Mensagem = "Arquivo gerado com sucesso!";
            return View("ChamarMacro");
        }

        public ActionResult ChamarMacro()
        {
            return View();
        }
    }
}