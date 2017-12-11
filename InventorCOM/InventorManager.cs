using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace InventorCOM
{
    class InventorManager
    {

        private Application app;

        public Application App {
            get { return app; }
        }

        public InventorManager()
        {
            this.app = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
            //Добавить возможность открывать Инвентор программно.
        }
    }
}
