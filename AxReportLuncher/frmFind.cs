using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion
{


    public partial class frmFind : Form
    {
        public enum Find
        {
            vender
            ,customer
            , Item
        }


        private TextBox _textBox ;
        private int _find;
       

        

        public frmFind()
        {
            InitializeComponent();
        //     _find = (int)Find;
         //   _textBox = AssignedValueTo;
        }

        
        /*
        public  void frmFind(Find Find,TextBox AssignedValueTo)
        {
           // InitializeComponent();
            _find = (int)Find;
            _textBox = AssignedValueTo;
        }
        */

       

        private void frmFind_Load(object sender, EventArgs e)
        {
            lblSearch.Text = "";
            txbSearch.Text = _textBox.Text;
           // gvList.DataSource =
        }

    }
}
