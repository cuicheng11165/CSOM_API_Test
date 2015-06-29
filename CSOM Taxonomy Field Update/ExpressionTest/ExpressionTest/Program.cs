using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionTest
{
    class Program
    {
        static void Main(string[] args)
        {
            dynamic test1 = new System.Dynamic.ExpandoObject();
            test1.DDD = "ddd";
            dynamic test = new { p1=1,p2=2,p3=3,p4=4,p5=5};

            //Func<int, int> express = x => x + 1;

            //var xVariable = Expression.Parameter(typeof(int), "x");
            //var cosntant1 = Expression.Constant(1);

            //var binaryExpress = Expression.Add(xVariable, cosntant1);

            //var lamdaExpress = Expression.Lambda(binaryExpress, xVariable);

            //var createdDelegate = lamdaExpress.Compile();

            //var result = createdDelegate.DynamicInvoke(1);



        }
    }
}
