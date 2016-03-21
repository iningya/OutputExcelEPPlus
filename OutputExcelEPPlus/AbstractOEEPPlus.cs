using OfficeOpenXml;
using System.Collections.Generic;

namespace OutputExcelEPPlus
{
    /// <summary>
    /// EPPluse を用いて Excel ファイル出力を行うベースクラス。
    /// </summary>
    public abstract class AbstractOEEPPlus
    {
        /// <summary>
        /// Excelファイルに対して行う処理
        /// </summary>
        /// <param name="ep">対象Excelファイル</param>
        protected delegate void ExcelPackageFunc(ExcelPackage ep);

        /// <summary>
        /// Excelファイルに対して行う処理のリスト。
        /// 処理実行順に追加すること。
        /// </summary>
        protected List<ExcelPackageFunc> ExcelPackageFuncList = new List<ExcelPackageFunc>();

        /// <summary>
        /// 処理リストを実行
        /// </summary>
        /// <param name="ep"></param>
        protected void ExecuteExcelPackageFuncList(ExcelPackage ep)
        {
            foreach (var func in this.ExcelPackageFuncList)
            {
                func(ep);
            }
        }

    }
}
