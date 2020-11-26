 grammar PoorExcel;

 /*
 * Parser Rules
 */

 compileUnit:expression EOF;

 expression:
 LPAREN expression RPAREN #ParenthesizedExpr
 | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
 | operatorToken=(INC|DEC)LPAREN expression RPAREN #IncDecExpr
 | expression operatorToken=(ADD | SUBSTRACT) expression #AdditiveExpr
 | operatorToken=(MOD|DIV)LPAREN expression DESP expression RPAREN#ModDivExpr
 | expression EXPONENT expression #ExponentialExpr
 | DECR expression #DecrExpr
 | INCRexpression #IncrPExpr
 | NUMBER #NumberExpr
 | IDENTIFIER #IdentifierExpr
 ;

 /*
 * Lexer Rules
 */

 NUMBER:INT('.'INT)?;
 IDENTIFIER:[a-zA-Z]+[1-9][0-9]+;
 INT : ('0'..'9')+;
 MULTIPLY:'*';
 DIVIDE:'/';
 ADD:'+';
 SUBSTRACT:'-';
MOD:'mod';
DIV:'div';
INC:'inc';
DEC:'dec';
EXPONENT:'^';
 LPAREN:'(';
 RPAREN:')';
 DESP:';'|',';
 WS:[\t\r\n]->channel(HIDDEN);

