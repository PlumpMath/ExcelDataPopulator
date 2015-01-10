using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class StringArrayConverter
    {
        internal static string[,] ConvertTo2DStringArrayStatically(IList<SampleObject> items)
        {   
            var stringArray = new string[items.Count, 5];
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                stringArray[i, 0] = item.A;
                stringArray[i, 1] = item.B;
                stringArray[i, 2] = item.C;
                stringArray[i, 3] = item.D;
                stringArray[i, 4] = item.E;
            }
            return stringArray;
        }

        internal static string[,] ConvertTo2DStringArrayWithReflection<T>(IList<T> items)
        {
            var props = typeof(T).GetProperties();
            var stringArray = new string[items.Count, props.Length];

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                for (int iProp = 0; iProp < props.Length; iProp++)
                {
                    stringArray[i, iProp] = props[iProp].GetValue(item).ToString();
                }
            }
            return stringArray;
        }

        internal static Func<IList<T>, string[,]> ConvertTo2DStringArrayWithExpressionTree<T>()
        {
            var t = typeof(T);
            var tProperties = t.GetProperties();

            var statements = new List<Expression>();

            var listInputParameter = Expression.Parameter(typeof(IList<T>));
            var itemPropertiesCount = Expression.Constant(tProperties.Length);
            var rowCountExpression = Expression.Property(listInputParameter, typeof(ICollection<T>), "Count");
            var arrayIndexVariable = Expression.Variable(typeof(int));
            var currentItem = Expression.Variable(t);
            
            statements.Add(Expression.Call(typeof(Console).GetMethod("WriteLine", new Type[] { typeof(string), typeof(string) })
                            , Expression.Constant("Converting {0} rows with expression tree")
                            , Expression.Call(rowCountExpression, typeof(object).GetMethod("ToString"))));

            // Initialize string array based on the items row count and the 
            var stringArrayVariable = Expression.Variable(typeof(string[,]));
            var newArrayExpression = Expression.NewArrayBounds(typeof(string), rowCountExpression, itemPropertiesCount);
            var initializeArray = Expression.Assign(stringArrayVariable, newArrayExpression);
            statements.Add(initializeArray);

            // Prepare item assignments here as it requires loop on the properties which can't be done inline
            var itemStatements = new List<Expression>();
            itemStatements.Add(Expression.Assign(currentItem, Expression.Property(listInputParameter, "Item", arrayIndexVariable)));
            for (int pIndex = 0; pIndex < tProperties.Length; pIndex++)
            {
                var propExpression = Expression.Property(currentItem, tProperties[pIndex]);
                itemStatements.Add(Expression.Assign(Expression.ArrayAccess(stringArrayVariable, new List<Expression> { arrayIndexVariable, Expression.Constant(pIndex) }), propExpression));
            }
            itemStatements.Add(Expression.PostIncrementAssign(arrayIndexVariable));

            // iterate the items
            var label = Expression.Label();
            var forLoopBody = Expression.Block(
                new[] { arrayIndexVariable }, // local variable
                Expression.Assign(arrayIndexVariable, Expression.Constant(0)), // initialize with 0
                Expression.Loop(
                    Expression.IfThenElse(
                        Expression.LessThan(arrayIndexVariable, rowCountExpression), // test
                        Expression.Block(new[] { currentItem }, itemStatements), // execute if true
                        Expression.Break(label)) // execute if false
                    , label));
            statements.Add(forLoopBody);

            // return statement
            statements.Add(stringArrayVariable); 

            var body = Expression.Block(stringArrayVariable.Type, new[] { stringArrayVariable }, statements.ToArray());
            var compiled = Expression.Lambda<Func<IList<T>, string[,]>>(body, listInputParameter).Compile();
            return compiled;
        }
    }
}
