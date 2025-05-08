grammar Expr;

prog        : ass ind dep stat;

ass         : ASS filePath ; // this is a file path, we will also parse through that file

filePath    : absolutePath | relativePath; // error with file paths when its path/file.txt '/' causes error

absolutePath: (WINDOWS_DRIVE | UNIX_ROOT) pathSegments?;

relativePath: pathSegments;

pathSegments: pathSegment ((UNIX_ROOT pathSegment)* | ( WINDOWS pathSegment)*) | ;
pathSegment: PATHSTRINGS | DOT | DOTDOT;



ind         : IND ind
            | TICK filePath ind
            |
            ;

dep         :
            | DEP dep
            | TICK filePath dep
            |
            ;

date        : DATE EQUAL num (comma num)*
            | DATE EQUAL num dash num
            ;
dash        : DASH;
comma       : COMMA;

DASH        : ' ' '-' ' '
            ;


COMMA       : ',';

num         : NUM;
DATE: 'date';
stat        :  date reg sum listall sort sorttype graphtype
            ;

graphtype   : GRAPH EQUAL graphtypes;
GRAPH : 'graph' | 'Graph';
graphtypes  : bar
            | line
            | pie
            | scatter
            ;
            bar : BAR;
            line : LINE;
            pie : PIE;
            scatter : SCATTER;


BAR : 'bar' | 'Bar';
LINE : 'line' | 'Line';
PIE : 'pie' | 'Pie';
SCATTER : 'scatter' | 'Scatter';



reg         : REGRESSION EQUAL regtypes
            |
            ;

regtypes    : linear
            | exponential
            | log
            | multiple
            | logistic
            | ridge
            ;

linear     : LINEAR
           ;

exponential : EXPONENTIAL
           ;

log        : LOG
           ;
multiple   : MULTIPLE
           ;
logistic  : LOGISTIC
           ;
ridge     : RIDGE
           ;
regression : REGRESSION
           ;


sum         : SUMMARY EQUAL TRUE
            | SUMMARY EQUAL FALSE
            |
            ;

listall      :   LISTALL EQUAL TRUE
             |  LISTALL EQUAL FALSE
             |
             ;

sort        : SORT EQUAL YEAR// these are always high to low unless sorttype is specified
            | SORT EQUAL // always
            |
            ;

sorttype     : SORTTYPE  EQUAL HIGHTOLOW // these should be changed to actions
             | SORTTYPE  EQUAL ALPHA // need to change these to actions
             | SORTTYPE EQUAL LOWTOHIGH
             | SORTTYPE  EQUAL BACKALPHA
             |
             ;

LINEAR : 'linear' | 'Linear';
EXPONENTIAL: 'exponential' | 'Exponential';
LOG     : 'log' | 'Log';
MULTIPLE    : 'multiple' | 'Multiple';
LOGISTIC    : 'logistic' | 'Logistic';
RIDGE       : 'ridge' | 'Ridge';
TRUE: 'true' | 'True';
FALSE: 'false' | 'False';
WINDOWS_DRIVE: [a-zA-Z] ':';
UNIX_ROOT: '/';
WINDOWS : '\\' ;
DOT: '.';
DOTDOT: '..';
SLASH: ('/' | '\\');
HIGHTOLOW     : 'HighToLow';
ALPHA       : 'Alpha';
LOWTOHIGH     : 'LowToHigh';
BACKALPHA   : 'BackAlpha';
YEAR        : 'year' ;
REGRESSION : 'regression'; // error with regression idk why
SORTTYPE : 'sorttype';
SORT    : 'sort';
LISTALL : 'listall';
SUMMARY : 'summary';
EQUAL  : '=';
COLON : ':';
ASS     : 'Association:';
IND     : 'Independent' ':';
DEP     : 'Dependent' ':';
TICK    : [A-Z]+;
NUM     : [1-9][0-9][0-9][0-9] ;
PATHSTRINGS : [a-zA-Z0-9_.-]+ ;
//LINEAR  : 'linear'; // these are all types of graphs
//EXPONENTIAL: ;
//LOG     : 'log';
//MULTIPLE    : 'multiple';
//LOGISTIC    : 'logistic';
//RIDGE       : 'ridge';

WSPACE: ' ' -> skip;
WS: [ \t\r\n]+ -> skip;
