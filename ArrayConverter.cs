using FastMember;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class ArrayConverter
    {
        internal static string[,] ConvertTo2DStringArrayStatically(IList<SampleObject> items)
        {
            var stringArray = new string[items.Count + 1, 9];
            // headers
            stringArray[0, 0] = "RowId";
            stringArray[0, 1] = "A";
            stringArray[0, 2] = "B";
            stringArray[0, 3] = "C";
            stringArray[0, 4] = "D";
            stringArray[0, 5] = "E";
            stringArray[0, 6] = "Qty";
            stringArray[0, 7] = "Price";
            stringArray[0, 8] = "Amount";

            // data
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var arrayIndex = i + 1;
                stringArray[arrayIndex, 0] = item.RowId.ToString();
                stringArray[arrayIndex, 1] = item.A;
                stringArray[arrayIndex, 2] = item.B;
                stringArray[arrayIndex, 3] = item.C;
                stringArray[arrayIndex, 4] = item.D;
                stringArray[arrayIndex, 5] = item.E;
                stringArray[arrayIndex, 6] = item.Qty.ToString();
                stringArray[arrayIndex, 7] = item.Price.ToString();
                stringArray[arrayIndex, 8] = item.Amount.ToString();
            }
            return stringArray;
        }

        internal static string[,] ConvertTo2DStringArrayWithReflection<T>(IList<T> items)
        {
            var props = typeof(T).GetProperties();
            var stringArray = new string[items.Count + 1, props.Length];
            // headers
            for (int iProp = 0; iProp < props.Length; iProp++)
            {
                stringArray[0, iProp] = props[iProp].Name;
            }

            // data
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var arrayIndex = i + 1;
                for (int iProp = 0; iProp < props.Length; iProp++)
                {
                    var val = props[iProp].GetValue(item);
                    if (val != null)
                        stringArray[arrayIndex, iProp] = val.ToString();
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
            var rowCountPlusHeaderExpression = Expression.Add(rowCountExpression, Expression.Constant(1));
            var itemIndexVariable = Expression.Variable(typeof(int));
            var arrayIndexVariable = Expression.Variable(typeof(int));
            var currentItem = Expression.Variable(t);

            statements.Add(Expression.Call(typeof(Console).GetMethod("WriteLine", new Type[] { typeof(string), typeof(string) })
                            , Expression.Constant("Converting {0} rows with expression tree")
                            , Expression.Call(rowCountExpression, typeof(object).GetMethod("ToString"))));

            // Initialize string array based on the items row count and the 
            var stringArrayVariable = Expression.Variable(typeof(string[,]));
            var newArrayExpression = Expression.NewArrayBounds(typeof(string), rowCountPlusHeaderExpression, itemPropertiesCount);
            var initializeArray = Expression.Assign(stringArrayVariable, newArrayExpression);
            statements.Add(initializeArray);

            // headers
            for (int pIndex = 0; pIndex < tProperties.Length; pIndex++)
            {
                statements.Add(Expression.Assign(Expression.ArrayAccess(stringArrayVariable, new List<Expression> { Expression.Constant(0), Expression.Constant(pIndex) }), Expression.Constant(tProperties[pIndex].Name)));
            }

            // Prepare item assignments here as it requires loop on the properties which can't be done inline
            var itemStatements = new List<Expression>();
            itemStatements.Add(Expression.Assign(currentItem, Expression.Property(listInputParameter, "Item", itemIndexVariable)));
            itemStatements.Add(Expression.Assign(arrayIndexVariable, Expression.Add(itemIndexVariable, Expression.Constant(1))));

            for (int pIndex = 0; pIndex < tProperties.Length; pIndex++)
            {
                var prop = tProperties[pIndex];
                var propExpression = Expression.Property(currentItem, prop);

                var callConvertExpression = Expression.Call(typeof(System.Convert).GetMethod("ToString", new Type[] { prop.PropertyType }), propExpression);
                // boxing is required here for excel to bind the data type correctly, converting all data to string will cause excel to treat numbers as text
                //var callConvertExpression = Expression.Convert(propExpression, typeof(object));

                itemStatements.Add(Expression.Assign(Expression.ArrayAccess(stringArrayVariable, new List<Expression> { arrayIndexVariable, Expression.Constant(pIndex) }), callConvertExpression));
            }
            itemStatements.Add(Expression.PostIncrementAssign(itemIndexVariable));

            // iterate the items
            var label = Expression.Label();
            var forLoopBody = Expression.Block(
                new[] { itemIndexVariable }, // local variable
                Expression.Assign(itemIndexVariable, Expression.Constant(0)), // initialize with 0
                Expression.Loop(
                    Expression.IfThenElse(
                        Expression.LessThan(itemIndexVariable, rowCountExpression), // test row count + header
                        Expression.Block(new[] { currentItem, arrayIndexVariable }, itemStatements), // execute if true
                        Expression.Break(label)) // execute if false
                    , label));
            statements.Add(forLoopBody);

            // return statement
            statements.Add(stringArrayVariable);

            var body = Expression.Block(stringArrayVariable.Type, new[] { stringArrayVariable }, statements.ToArray());
            var compiled = Expression.Lambda<Func<IList<T>, string[,]>>(body, listInputParameter).Compile();
            return compiled;
        }

        internal static object[,] ConvertTo2DObjectArrayStatically(IList<SampleObject> items)
        {
            var stringArray = new object[items.Count + 1, 9];
            // headers
            stringArray[0, 0] = "RowId";
            stringArray[0, 1] = "A";
            stringArray[0, 2] = "B";
            stringArray[0, 3] = "C";
            stringArray[0, 4] = "D";
            stringArray[0, 5] = "E";
            stringArray[0, 6] = "Qty";
            stringArray[0, 7] = "Price";
            stringArray[0, 8] = "Amount";

            // data
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var arrayIndex = i + 1;
                stringArray[arrayIndex, 0] = item.RowId;
                stringArray[arrayIndex, 1] = item.A;
                stringArray[arrayIndex, 2] = item.B;
                stringArray[arrayIndex, 3] = item.C;
                stringArray[arrayIndex, 4] = item.D;
                stringArray[arrayIndex, 5] = item.E;
                stringArray[arrayIndex, 6] = item.Qty;
                stringArray[arrayIndex, 7] = item.Price;
                stringArray[arrayIndex, 8] = item.Amount;
            }
            return stringArray;
        }

        internal static object[,] ConvertTo2DObjectArrayWithReflection<T>(IList<T> items)
        {
            var props = typeof(T).GetProperties();
            var stringArray = new object[items.Count + 1, props.Length];
            // headers
            for (int iProp = 0; iProp < props.Length; iProp++)
            {
                stringArray[0, iProp] = props[iProp].Name;
            }

            // data
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var arrayIndex = i + 1;
                for (int iProp = 0; iProp < props.Length; iProp++)
                {
                    var val = props[iProp].GetValue(item);
                    if (val != null)
                        stringArray[arrayIndex, iProp] = val;
                }
            }
            return stringArray;
        }

        internal static Func<IList<T>, object[,]> ConvertTo2DObjectArrayWithExpressionTree<T>()
        {
            var t = typeof(T);
            var tProperties = t.GetProperties();

            var statements = new List<Expression>();

            var listInputParameter = Expression.Parameter(typeof(IList<T>));
            var itemPropertiesCount = Expression.Constant(tProperties.Length);
            var rowCountExpression = Expression.Property(listInputParameter, typeof(ICollection<T>), "Count");
            var rowCountPlusHeaderExpression = Expression.Add(rowCountExpression, Expression.Constant(1));
            var itemIndexVariable = Expression.Variable(typeof(int));
            var arrayIndexVariable = Expression.Variable(typeof(int));
            var currentItem = Expression.Variable(t);

            //statements.Add(Expression.Call(typeof(Debug).GetMethod("WriteLine", new Type[] { typeof(string), typeof(string) })
            //                , Expression.Constant("Converting {0} rows with expression tree")
            //                , Expression.Call(rowCountExpression, typeof(object).GetMethod("ToString"))));

            // Initialize string array based on the items row count and the 
            var stringArrayVariable = Expression.Variable(typeof(object[,]));
            var newArrayExpression = Expression.NewArrayBounds(typeof(object), rowCountPlusHeaderExpression, itemPropertiesCount);
            var initializeArray = Expression.Assign(stringArrayVariable, newArrayExpression);
            statements.Add(initializeArray);

            // headers
            for (int pIndex = 0; pIndex < tProperties.Length; pIndex++)
            {
                statements.Add(Expression.Assign(Expression.ArrayAccess(stringArrayVariable, new List<Expression> { Expression.Constant(0), Expression.Constant(pIndex) }), Expression.Constant(tProperties[pIndex].Name)));
            }

            // Prepare item assignments here as it requires loop on the properties which can't be done inline
            var itemStatements = new List<Expression>();
            itemStatements.Add(Expression.Assign(currentItem, Expression.Property(listInputParameter, "Item", itemIndexVariable)));
            itemStatements.Add(Expression.Assign(arrayIndexVariable, Expression.Add(itemIndexVariable, Expression.Constant(1))));

            for (int pIndex = 0; pIndex < tProperties.Length; pIndex++)
            {
                var prop = tProperties[pIndex];
                var propExpression = Expression.Property(currentItem, prop);

                //var callConvertExpression = Expression.Call(typeof(System.Convert).GetMethod("ToString", new Type[] { prop.PropertyType }), propExpression);
                // boxing is required here for excel to bind the data type correctly, converting all data to string will cause excel to treat numbers as text
                var callConvertExpression = Expression.Convert(propExpression, typeof(object));

                itemStatements.Add(Expression.Assign(Expression.ArrayAccess(stringArrayVariable, new List<Expression> { arrayIndexVariable, Expression.Constant(pIndex) }), callConvertExpression));
            }
            itemStatements.Add(Expression.PostIncrementAssign(itemIndexVariable));

            // iterate the items
            var label = Expression.Label();
            var forLoopBody = Expression.Block(
                new[] { itemIndexVariable }, // local variable
                Expression.Assign(itemIndexVariable, Expression.Constant(0)), // initialize with 0
                Expression.Loop(
                    Expression.IfThenElse(
                        Expression.LessThan(itemIndexVariable, rowCountExpression), // test row count + header
                        Expression.Block(new[] { currentItem, arrayIndexVariable }, itemStatements), // execute if true
                        Expression.Break(label)) // execute if false
                    , label));
            statements.Add(forLoopBody);

            // return statement
            statements.Add(stringArrayVariable);

            var body = Expression.Block(stringArrayVariable.Type, new[] { stringArrayVariable }, statements.ToArray());
            var compiled = Expression.Lambda<Func<IList<T>, object[,]>>(body, listInputParameter).Compile();
            return compiled;
        }

        internal static object[,] ConvertTo2DObjectArrayWithFastMember<T>(IList<T> items)
        {
            //Debug.WriteLine("Fast member - start");
            var t = typeof(T);
            var props = t.GetProperties();
            var stringArray = new object[items.Count + 1, props.Length];
            var propNames = props.Select(x => x.Name).ToArray();
            // headers
            for (int iProp = 0; iProp < props.Length; iProp++)
            {
                stringArray[0, iProp] = props[iProp].Name;
            }

            var accessor = TypeAccessor.Create(t, false);

            //Debug.WriteLine("Fast member - write data");
            // data
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var arrayIndex = i + 1;
                for (int iProp = 0; iProp < props.Length; iProp++)
                {
                    var val = accessor[item, propNames[iProp]];
                    if (val != null)
                        stringArray[arrayIndex, iProp] = val;
                }
            }
            return stringArray;
        }
    }
}
