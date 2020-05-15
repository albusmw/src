using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.Windows.Forms;
using DynamicTypeDescriptor;

using Dyn = DynamicTypeDescriptor;
using Scm = System.ComponentModel;

namespace CustomTypeDescriptorApp
{
  public partial class Form2 : Form
  {
    public Form2()
    {
      InitializeComponent( );
      Dyn.ResourceAttribute ra = new Dyn.ResourceAttribute( );
      ra.AssemblyFullName = this.GetType( ).Assembly.FullName;
      ra.KeyPrefix = "BuiltinBool_";
      ra.BaseName = "CustomTypeDescriptorApp.Properties.Resources";
      Scm.TypeConverterAttribute tca = new TypeConverterAttribute(typeof(Dyn.BooleanConverter));

      // setting resource for boolean in global scope
      Scm.TypeDescriptor.AddAttributes(typeof(Boolean), ra, tca);
    }

    /// <summary>
    /// This enum should be created with maximum number of possible values.
    /// And it should placed in a generic location so it can be reused.
    /// </summary>
    [Scm.Editor(typeof(Dyn.StandardValueEditor), typeof(UITypeEditor))]
    [Scm.TypeConverter(typeof(Dyn.EnumConverter))]
    [Flags( )]
    internal enum PseudoEnum : ulong
    {
      Enum_0 = 0,
      Enum_1 = 1,
      Enum_2 = 2,
      Enum_3 = 4,
      Enum_4 = 8,
      Enum_5 = 16,
      Enum_6 = 32,
      Enum_7 = 64,
      Enum_8 = 128,
      Enum_9 = 256,
      Enum_10 = 512,
      Enum_11 = 1024,
      Enum_12 = 2048,
      Enum_13 = 4096,
      Enum_14 = 8192,
      Enum_15 = 16384,
      Enum_16 = 32768,
      Enum_17 = 65536,
      Enum_18 = 131072,
      Enum_19 = 262144,
      Enum_20 = 524288,
      Enum_21 = 1048576,
      Enum_22 = 2097152,
      Enum_23 = 4194304,
      Enum_24 = 8388608,
      Enum_25 = 16777216,
      Enum_26 = 33554432,
      Enum_27 = 67108864,
      Enum_28 = 134217728,
      Enum_29 = 268435456,
      Enum_30 = 536870912,
      Enum_31 = 1073741824,
    }

    private void Form1_Load( System.Object sender, System.EventArgs e )
    {
      this.PropertyGrid1.SelectedObject = new Phone( );
      PropertyGrid1.PropertySort = PropertySort.Categorized;
      PropertyGrid1.Refresh( );
    }

    [Dyn.DisplayName("Phone Number")]
    internal class Phone
    {
      public Phone()
      {
        Dyn.TypeDescriptor.IntallTypeDescriptor(this);
        Dyn.TypeDescriptor.GetTypeDescriptor(this).GetProperties( ).Clear( );

        // first property (on-the-fly)
        Dyn.PropertyDescriptor pdA = new Dyn.PropertyDescriptor(this.GetType( ), "FlagPropertyA", typeof(PseudoEnum), PseudoEnum.Enum_6);
        pdA.Attributes.Add(new Scm.BrowsableAttribute(true));
        pdA.Attributes.Add(new Scm.DescriptionAttribute("Produced from a database Enum"));
        pdA.Attributes.Add(new Scm.DefaultValueAttribute(PseudoEnum.Enum_6));
        pdA.Attributes.Add(new Dyn.ExclusiveStandardValuesAttribute(true), true);
        pdA.AddValueChanged(this, new EventHandler(this.OnTheFlyPropertyChanged));

        Dyn.EnumPropertyDescriptor enuPdA = new Dyn.EnumPropertyDescriptor(pdA);
        SynchronizeEnums(enuPdA.StandardValues, GetEnumFromDatabaseA( ));
        Dyn.TypeDescriptor.GetTypeDescriptor(this).GetProperties( ).Add(enuPdA);

        // second property (on-the-fly)
        Dyn.PropertyDescriptor pdB = new Dyn.PropertyDescriptor(this.GetType( ), "FlagPropertyB", typeof(PseudoEnum), PseudoEnum.Enum_4);
        pdB.Attributes.Add(new Scm.BrowsableAttribute(true));
        pdB.Attributes.Add(new Scm.DescriptionAttribute("Produced from a database Enum"));
        pdB.Attributes.Add(new Scm.DefaultValueAttribute(PseudoEnum.Enum_4));
        pdB.Attributes.Add(new Dyn.ExclusiveStandardValuesAttribute(true), true);
        pdB.AddValueChanged(this, new EventHandler(this.OnTheFlyPropertyChanged));

        Dyn.EnumPropertyDescriptor enuPdB = new Dyn.EnumPropertyDescriptor(pdB);
        SynchronizeEnums(enuPdB.StandardValues, GetEnumFromDatabaseB( ));
        Dyn.TypeDescriptor.GetTypeDescriptor(this).GetProperties( ).Add(enuPdB);
      }

      private void OnTheFlyPropertyChanged( object sender, EventArgs e )
      {
        // Save the changes
      }

      private IList<StandardValue> GetEnumFromDatabaseA()
      {
        List<StandardValue> list = new List<StandardValue>( );
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 0), "None", "Nothing selected"));
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 32), "Name of 32", "Selection 32"));
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 512), "Name of 512", "Selection 512"));

        // NO, can't do this. 9 is not one of the value in PseodoEnum
        //list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 9), "Name of 9", "Selection 9"));
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 64), "Name of 64", "Selection 64"));
        return list;
      }

      private IList<StandardValue> GetEnumFromDatabaseB()
      {
        List<StandardValue> list = new List<StandardValue>( );
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 8), "Name of 8", "Selection 8"));
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 2), "Name of 2", "Selection 2"));
        list.Add(new StandardValue(Enum.ToObject(typeof(PseudoEnum), 524288), "Name of 524288", "Selection 524288"));
        return list;
      }

      /// <summary>
      /// This methods shows enums of property that are in the database list, and hide others
      /// </summary>
      /// <param name="fromProperty">all possible values of the enum</param>
      /// <param name="fromDatabase">a subset of the values from the database</param>
      /// <param name="enumType">type of the enum</param>
      private void SynchronizeEnums( IList<StandardValue> fromProperty, IList<StandardValue> fromDatabase )
      {
        foreach (StandardValue svProp in fromProperty)
        {
          StandardValue svFound = null;
          foreach (StandardValue svDB in fromDatabase)
          {
            if (svProp.Value.Equals(svDB.Value))
            {
              svFound = svDB;
              break;
            }
          }
          if (svFound != null)
          {
            svProp.Description = svFound.Description;
            svProp.DisplayName = svFound.DisplayName;
            svProp.Visible = true;
            svProp.Enabled = true;
          }
          else
          {
            svProp.Visible = false;
          }
        }
      }
    }

    private void button1_Click( object sender, EventArgs e )
    {
    }
  }
}