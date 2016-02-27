#include "threademp.h"
#include "exportxml.h"

ThreadEmp::ThreadEmp()
{
    stopped = false;
}

void ThreadEmp::stop()
{

}

void ThreadEmp::run()
{
    ExportXML exportXML;
    exportXML.openImport();
}
