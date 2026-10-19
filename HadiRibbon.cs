
using System;
using ExcelDna.Integration.CustomUI;

namespace Hadi.Excel
{
    public class HadiRibbon : ExcelRibbon
    {
        private static IRibbonUI _ribbon;

        public override string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='hadiTab' label='Hadi Tools'>

        <group id='maxMinGroup' label='Max / Min'>
          <toggleButton id='btnMax' label='Max' size='large'
                        onAction='OnMaxClick'
                        getPressed='GetMaxPressed'
                        imageMso='ArrowUp'
                        screentip='Highlight the maximum point'
                        supertip='Toggle a red ring on the maximum point of the selected series. Updates live as data changes.' />

          <toggleButton id='btnMin' label='Min' size='large'
                        onAction='OnMinClick'
                        getPressed='GetMinPressed'
                        imageMso='ArrowDown'
                        screentip='Highlight the minimum point'
                        supertip='Toggle a blue ring on the minimum point of the selected series. Updates live as data changes.' />
        </group>

        <group id='actionsGroup' label='Actions'>
          <button id='btnClearAll' label='Clear All' size='large'
                  onAction='OnClearAllClick'
                  imageMso='Delete'
                  screentip='Clear highlights'
                  supertip='If a chart is selected, clears all highlights on that chart. Otherwise asks whether to clear highlights on every chart in the active sheet.' />
        </group>

        <group id='waterCadGroup' label='WaterCAD'>
          <button id='btnReadExcelTable' label='Read Table' size='large'
                  onAction='OnReadExcelTableClick'
                  imageMso='TableSelect'
                  screentip='Select and read an Excel table'
                  supertip='Select a range of cells and preview its headers and values.' />
          <button id='btnReadJunctions' label='Read Junctions' size='large'
                  onAction='OnReadJunctionsClick'
                  imageMso='ShapeOval'
                  screentip='Read and map junction data'
                  supertip='Read junction labels, elevations, and optional X/Y coordinates from Excel.' />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Log.Write("Ribbon loaded.");
        }

        public static void InvalidateControls()
        {
            try { _ribbon?.Invalidate(); } catch { }
        }

        public void OnMaxClick(IRibbonControl control, bool pressed)
        {
            MaxMin.Toggle("Max");
            InvalidateControls();
        }

        public void OnMinClick(IRibbonControl control, bool pressed)
        {
            MaxMin.Toggle("Min");
            InvalidateControls();
        }

        public void OnClearAllClick(IRibbonControl control)
        {
            MaxMin.ClearAll();
            InvalidateControls();
        }

        public bool GetMaxPressed(IRibbonControl control)
            => MaxMin.IsActive("Max");

        public bool GetMinPressed(IRibbonControl control)
            => MaxMin.IsActive("Min");

        public void OnReadExcelTableClick(IRibbonControl control)
        {
            ExcelTableReader.ReadAndPreview();
        }

        public void OnReadJunctionsClick(IRibbonControl control)
        {
            ExcelTableReader.ReadJunctionsAndPreview();
        }
    }
}
