using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using Dyn = DynamicTypeDescriptor;
using System.Threading;

namespace DynamicTypeDescriptor
{
  internal partial class StandardValueEditorUI : UserControl
  {
    private class TagItem
    {
      public bool SetByCode = false;
      public StandardValue Item = null;

      public TagItem( StandardValue item )
      {
        Item = item;
      }
    }

    private object m_Value = null;
    private IWindowsFormsEditorService m_editorService = null;
    private ITypeDescriptorContext m_context = null;
    private bool m_bFlag = false;

    public StandardValueEditorUI()
    {
      InitializeComponent( );
    }

    public void SetData( ITypeDescriptorContext context, IWindowsFormsEditorService editorService, object value )
    {
      Debug.Assert(editorService != null);
      Debug.Assert(context.PropertyDescriptor != null);
      Debug.Assert(editorService != null);

      m_editorService = editorService;
      m_context = context;
      m_Value = value;

      m_bFlag = (context.PropertyDescriptor.PropertyType.GetCustomAttributes(typeof(FlagsAttribute), false).Length > 0);

      listViewSvc.Items.Clear( );
      listViewSvc.CheckBoxes = m_bFlag;

      TypeConverter tc = context.PropertyDescriptor.Converter;
      Debug.Assert(tc.GetStandardValuesSupported(context));
      Debug.Assert((tc.CanConvertTo(context, typeof(StandardValue))), "To use this editor, TypeConverter must be able to convert to StandardValue");

      TypeConverter.StandardValuesCollection svc = tc.GetStandardValues(context);

      // create list view items for the visible Enum items
      foreach (object obj in svc)
      {
        ListViewItem lvi = new ListViewItem( );
        StandardValue sv = tc.ConvertTo(context, Thread.CurrentThread.CurrentUICulture, obj, typeof(StandardValue)) as StandardValue;
        if (sv != null && sv.Visible)
        {
          lvi.Text = sv.DisplayName;
          lvi.ForeColor = (sv.Enabled == true ? lvi.ForeColor : Color.FromKnownColor(KnownColor.GrayText));
          lvi.Tag = new TagItem(sv);
          listViewSvc.Items.Add(lvi);
        }
      }

      UpdateCheckState( );

      // make initial selection
      if (m_context.PropertyDescriptor.PropertyType.IsEnum && m_bFlag)
      {
        // select the first checked one
        foreach (ListViewItem lvi in listViewSvc.CheckedItems)
        {
          lvi.Selected = true;
          lvi.EnsureVisible( );
          listViewSvc.FocusedItem = lvi;
          break;
        }
      }
      else
      {
        foreach (ListViewItem lvi in listViewSvc.Items)
        {
          TagItem ti = lvi.Tag as TagItem;
          if (ti.Item.Value.Equals(m_Value))
          {
            lvi.Selected = true;
            lvi.EnsureVisible( );
            listViewSvc.FocusedItem = lvi;
            break;
          }
        }
      }
    }

    public object GetValue()
    {
      if (m_context.PropertyDescriptor.PropertyType.IsEnum)
      {
        return Enum.ToObject(m_context.PropertyDescriptor.PropertyType, m_Value);
      }
      return m_Value;
    }

    private void listViewSvc_ItemCheck( object sender, ItemCheckEventArgs e )
    {
      TagItem ti = listViewSvc.Items[e.Index].Tag as TagItem;

      if (ti.SetByCode)
      {
        ti.SetByCode = false;
        return;
      }
      if (!ti.Item.Enabled)
      {
        e.NewValue = e.CurrentValue;
        return;
      }

      if (e.NewValue == CheckState.Checked)
      {
        bool bOk = EnumUtil.TurnOnBits(ref m_Value, ti.Item.Value);
      }
      else
      {
        bool bOk = EnumUtil.TurnOffBits(ref m_Value, ti.Item.Value);
      }

      e.NewValue = e.CurrentValue;
      UpdateCheckState( ); // this will change the check box on the list view item
    }

    private void listViewSvc_SelectedIndexChanged( object sender, EventArgs e )
    {
      if (listViewSvc.SelectedItems.Count > 0)
      {
        ListView listView = (ListView)sender;

        TagItem ti = listView.SelectedItems[0].Tag as TagItem;

        lblDispName.Text = ti.Item.DisplayName;
        lblDesc.Text = ti.Item.Description;

        if (!m_bFlag && ti.Item.Enabled)
        {
          m_Value = ti.Item.Value;
        }
      }
    }

    private void listViewSvc_MouseDoubleClick( object sender, MouseEventArgs e )
    {
      m_editorService.CloseDropDown( );
    }

    private void listViewSvc_SizeChanged( object sender, EventArgs e )
    {
      listViewSvc.Columns[0].Width = listViewSvc.Width - 20;
      listViewSvc.Invalidate( );
      lblDesc.Invalidate( );
      this.Invalidate( );
    }

    private void UpdateCheckState()
    {
      if (m_context.PropertyDescriptor.PropertyType.IsEnum && m_bFlag)
      {
        foreach (ListViewItem lvi in listViewSvc.Items)
        {
          TagItem ti = lvi.Tag as TagItem;
          bool bitExist = EnumUtil.IsBitsOn(m_Value, ti.Item.Value);
          if (lvi.Checked != bitExist)
          {
            ti.SetByCode = true;
            lvi.Checked = bitExist;
          }
        }
      }
    }
  }
}