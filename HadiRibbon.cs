using ExcelDna.Integration.CustomUI;

namespace Hadi.Excel
{
    public class HadiRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='hadiTab' label='Hadi Tools'>
        <group id='maxMinGroup' label='Max / Min'>
          <button id='btnMax' label='Max' size='normal'
                  onAction='OnMaxClick' imageMso='ArrowUp' />
          <button id='btnMin' label='Min' size='normal'
                  onAction='OnMinClick' imageMso='ArrowDown' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnMaxClick(IRibbonControl control) => MaxMin.Toggle("Max");
        public void OnMinClick(IRibbonControl control) => MaxMin.Toggle("Min");
    }
}