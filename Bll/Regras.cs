using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bolinha.Modelos;
using Bolinha.DAL;

namespace Bolinha.Bll
{
    public class Regras
    {
        //******************************************************************************************************************************
        //Procedure para alternar cor de linha do DatagridView
        public void CorDataGrid(System.Windows.Forms.DataGridView grid)
        {
            int n = grid.RowCount;
            for (int i = 0; i < n; i++)
            {
                if (i % 2 == 0)
                {
                    //grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Silver;
                    grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.BurlyWood;
                }
                else
                {
                    grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                }

                //Bitmap grayScale = new Bitmap(image.Width, image.Height);

            }
        }

        //******************************************************************************************************************************
    }
}
