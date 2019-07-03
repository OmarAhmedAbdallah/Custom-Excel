using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateExcelSheet.Models
{
    class Employee
    {
        public int EmployeeId { set; get; }
        public string EmployeeFirstName { set; get; }
        public string EmployeeLastName { set; get; }
        public string EmployeeFloor { set; get; }

        public double EmployeeBonus { set; get; }

        public Employee(int employeeId, string employeeFirstName, string employeeLastName,  string employeeFloor,double employeeBonus)
        {
            EmployeeId = employeeId;
            EmployeeFirstName = employeeFirstName;
            EmployeeLastName = employeeLastName;
            EmployeeFloor = employeeFloor;
            EmployeeBonus = employeeBonus;
        }

        public static List<Employee> GetEmployees()
        {
            return new List<Employee>()
                {
                    new Employee(1,"Ahmed","Zaki","1st",1000.0),
                    new Employee(2,"Mohamed","Hegazi","2nd",1000.0),
                    new Employee(3,"Yousef","Saad","1st",1000.0),
                    new Employee(4,"Mai","Gamel","1st",1000.0),
                    new Employee(5,"Sara","Badr","2nd",1000.0),
                };
        }
    }
}
