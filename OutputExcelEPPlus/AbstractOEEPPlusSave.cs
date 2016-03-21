using OfficeOpenXml;
using System.IO;

namespace OutputExcelEPPlus
{
    /// <summary>
    /// EPPluse を用いて Excel ファイル出力（上書き）を行うベースクラス。
    /// </summary>
    public abstract class AbstractOEEPPlusSave : AbstractOEEPPlus
    {
        /// <summary>
        /// 上書き対象のファイル
        /// </summary>
        private FileInfo Fi;

        /// <summary>
        /// コンストラクタで対象ファイルを設定。
        /// </summary>
        /// <param name="filePath">上書き対象のファイル</param>
        protected AbstractOEEPPlusSave(string filePath)
        {
            this.Fi = new FileInfo(filePath);
        }

        /// <summary>
        /// ファイルに対する処理を実行し上書き保存
        /// </summary>
        protected void Execute()
        {
            using (var ep = new ExcelPackage(this.Fi))
            {
                this.ExecuteExcelPackageFuncList(ep);
                ep.Save();
            }
        }

    }
}
