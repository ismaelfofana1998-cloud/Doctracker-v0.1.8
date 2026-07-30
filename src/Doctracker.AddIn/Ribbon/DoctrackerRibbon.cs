using System;
using System.Runtime.InteropServices;
using Doctracker.AddIn.UI;
using Doctracker.Core.Models;
using Microsoft.Office.Core;

namespace Doctracker.AddIn.Ribbon
{
    [ComVisible(true)]
    public sealed class DoctrackerRibbon : IRibbonExtensibility
    {
        private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='DoctrackerTab' label='Doctracker'>
        <group id='ProjectGroup' label='Dossier'>
          <button id='OpenPane' label='Ouvrir Doctracker' size='large' onAction='OpenPane_OnAction'/>
          <button id='ImportDocuments' label='Ajouter des pièces' onAction='ImportDocuments_OnAction'/>
        </group>
        <group id='SnipGroup' label='Snips'>
          <button id='TextSnip' label='Texte' onAction='TextSnip_OnAction'/>
          <button id='NumberSnip' label='Nombre' onAction='NumberSnip_OnAction'/>
          <button id='DateSnip' label='Date' onAction='DateSnip_OnAction'/>
          <button id='SumSnip' label='Somme' onAction='SumSnip_OnAction'/>
          <button id='TableSnip' label='Tableau' onAction='TableSnip_OnAction'/>
        </group>
        <group id='MatchingGroup' label='Contrôle'>
          <button id='Match' label='Document Matching' size='large' onAction='Match_OnAction'/>
          <button id='OpenProof' label='Ouvrir la preuve' onAction='OpenProof_OnAction'/>
          <button id='ReviewProof' label='Revoir' onAction='ReviewProof_OnAction'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        public void OnLoad(IRibbonUI ribbonUi)
        {
            ribbon = ribbonUi;
        }

        public void OpenPane_OnAction(IRibbonControl control) => Controller.Toggle();
        public void ImportDocuments_OnAction(IRibbonControl control) => Controller.ImportDocuments();
        public void TextSnip_OnAction(IRibbonControl control) => Controller.CreateSnip(SnipType.Text);
        public void NumberSnip_OnAction(IRibbonControl control) => Controller.CreateSnip(SnipType.Number);
        public void DateSnip_OnAction(IRibbonControl control) => Controller.CreateSnip(SnipType.Date);
        public void SumSnip_OnAction(IRibbonControl control) => Controller.CreateSnip(SnipType.Sum);
        public void TableSnip_OnAction(IRibbonControl control) => Controller.CreateSnip(SnipType.Table);
        public void Match_OnAction(IRibbonControl control) => Controller.MatchSelection();
        public void OpenProof_OnAction(IRibbonControl control) => Controller.NavigateFromSelection();
        public void ReviewProof_OnAction(IRibbonControl control) => Controller.ReviewSelection();

        private static PaneController Controller
        {
            get
            {
                var controller = Globals.ThisAddIn?.Controller;
                if (controller == null)
                    throw new InvalidOperationException("Doctracker is not initialized.");
                return controller;
            }
        }
    }
}
