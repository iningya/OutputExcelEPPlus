using OfficeOpenXml;
using System.IO;

namespace OutputExcelEPPlus
{
    /// <summary>
    /// EPPluse を用いて Excel ファイル出力（別名保存）を行うベースクラス。
    /// </summary>
    class AbstractOEEPPlusSaveAs : AbstractOEEPPlus
    {
        /// <summary>
        /// ベースとなるファイル
        /// </summary>
        private FileInfo TemplateFi;

        /// <summary>
        /// 出力するファイル
        /// </summary>
        private FileInfo OutputFi;

        /// <summary>
        ///// コンストラクタで対象ファイルを設定。
        /// </summary>
        /// <param name="templateFilePath">ベースとなるファイル</param>
        /// <param name="outputFilePath">出力ファイル</param>
        public AbstractOEEPPlusSaveAs(string templateFilePath, string outputFilePath)
        {
            this.TemplateFi = new FileInfo(templateFilePath);
            this.OutputFi = new FileInfo(outputFilePath);
        }

        /// <summary>
        /// ファイルに対する処理を実行しファイル出力j
        /// </summary>
        protected void Execute()
        {
            using (var ep = new ExcelPackage(this.TemplateFi))
            {
                this.ExecuteExcelPackageFuncList(ep);
                ep.SaveAs(this.OutputFi);
            }
        }

    }
}
