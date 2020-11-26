using Antlr4.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PoorExcel
{
    class PoorExcelVisitor : PoorExcelBaseVisitor<double>
    {
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();
        public override double VisitCompileUnit(PoorExcelParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitNumberExpr(PoorExcelParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }
        public override double VisitIdentifierExpr(PoorExcelParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
                return value;
            else
                return 0.0;
        }
        public override double VisitParenthesizedExpr(PoorExcelParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitAdditiveExpr([NotNull] PoorExcelParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == PoorExcelLexer.ADD)
            {
                Debug.WriteLine("{0}+{1}", left, right);
                return left + right;
            }
            else
            {
                Debug.WriteLine("{0}-{1}", left, right);
                return left - right;
            }
        }
        public override double VisitMultiplicativeExpr([NotNull] PoorExcelParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == PoorExcelLexer.MULTIPLY)
            {
                Debug.WriteLine("{0}*{1}", left, right);
                return left * right;
            }
            else
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }

        public override double VisitIncDecExpr([NotNull] PoorExcelParser.IncDecExprContext context)
        {
            var number = WalkLeft(context);
            if (context.operatorToken.Type == PoorExcelLexer.INC)
            {
                Debug.WriteLine("inc({0})", number);
                return number + 1;
            }
            else
            {
                Debug.WriteLine("dec({0})", number);
                return number - 1;
            }
        }
        public override double VisitModDivExpr([NotNull] PoorExcelParser.ModDivExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == PoorExcelLexer.MOD)
                return left % right;
            else
                return (int)left / (int)right;
        }
        public override double VisitExponentialExpr([NotNull] PoorExcelParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0}^{1}", left, right);
            return System.Math.Pow(left, right);
        }
        private double WalkLeft(PoorExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<PoorExcelParser.ExpressionContext>(0));
        }

        private double WalkRight(PoorExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<PoorExcelParser.ExpressionContext>(1));
        }
       

    }
    
   
}
