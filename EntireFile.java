import com.example.ExcelKeywordExtractor;
import org.antlr.v4.runtime.*;

import java.io.Console;
import java.util.ArrayList;
import java.util.List;

import java.io.PrintStream;
import java.util.Vector;

class LangListener extends ExprBaseListener {
    String AssfilePath = "";

    String depFilePath = "";

    String indFilePath = "";
    String graphType = "";
    String regType = "";
    Vector<Integer> years = new Vector<>();
    String dateString = "";

    String sortType = "";

    boolean summary = false;
    // @Override
    public void exitProg(ExprParser.ProgContext ctx) {
        //System.out.println("exited Prog");
    }
    // @Override
    @Override
    public void exitAss(ExprParser.AssContext ctx) {
        AssfilePath = ctx.filePath().getText();
        System.out.println("Association File path: " + AssfilePath);
    }
    @Override
    public void exitGraphtype(ExprParser.GraphtypeContext ctx) {
       //graphType = ctx.getText();
       // System.out.println("Graph Type: " + graphType);
    }
    @Override
    public void exitBar(ExprParser.BarContext ctx) {
        if (ctx.BAR() != null) {
            graphType = ctx.BAR().getText();
            System.out.println("Graph Type: " + graphType);
        }
    }
    @Override
    public void exitLine(ExprParser.LineContext ctx) {
        if (ctx.LINE() != null) {
            graphType = ctx.LINE().getText();
            System.out.println("Graph Type: " + graphType);
        }
    }
    @Override
    public void exitPie(ExprParser.PieContext ctx) {
        if (ctx.PIE() != null) {
            graphType = ctx.PIE().getText();
            System.out.println("Graph Type: " + graphType);
        }
    }
    @Override
    public void exitScatter(ExprParser.ScatterContext ctx) {
        if (ctx.SCATTER() != null) {
            graphType = ctx.SCATTER().getText();
            System.out.println("Graph Type: " + graphType);
        }
    }
    @Override
    public void exitInd(ExprParser.IndContext ctx) {
        if (ctx.filePath() != null) {
            if (!indFilePath.isEmpty()) {
                indFilePath += ", ";
            }
            indFilePath = ctx.filePath().getText();
            System.out.println("Independent File path: " + indFilePath);
        }
    }

    @Override
    public void exitDep(ExprParser.DepContext ctx) {
        if (ctx.filePath() != null) {
            if (!depFilePath.isEmpty()) {
                depFilePath += ", ";
            }
            depFilePath += ctx.filePath().getText();
            System.out.println("Dependent File path: " + depFilePath);
        }
    }

    @Override
    public void exitDate(ExprParser.DateContext ctx){



        if (ctx.comma() != null && !ctx.comma().isEmpty()){
            for (ExprParser.NumContext numCtx : ctx.num()) {
                int year = Integer.parseInt(numCtx.getText());
                years.add(year);
            }
        }
        if (ctx.dash() != null && !ctx.dash().isEmpty()) {
//            System.out.println("EXAMPLE: " + ctx.num(0).getText());
            for (int i = (Integer.parseInt(ctx.num(0).getText())); i <= (Integer.parseInt(ctx.num(1).getText())); i++) {
//                System.out.println("EXAMPLE: " + i);
                years.add(i);
            }
        }
        System.out.println("Years: " + years);

    }


    @Override public void exitLinear (ExprParser.LinearContext ctx) {
        if (ctx.LINEAR() != null) {
            regType = ctx.LINEAR().getText();
            System.out.println("RegType File path: " + regType);
        }
    }

    @Override public void exitExponential (ExprParser.ExponentialContext ctx) {
        if (ctx.EXPONENTIAL() != null) {
             regType = ctx.EXPONENTIAL().getText();
            System.out.println("RegType File path: " + regType);
        }
    }

    @Override public void exitLog(ExprParser.LogContext ctx) {
        if (ctx.LOG() != null) {
            regType = ctx.LOG().getText();
            System.out.println("RegType File path: " + regType);
        }
    }

    @Override public void exitMultiple(ExprParser.MultipleContext ctx) {
        if (ctx.MULTIPLE() != null) {
            regType = ctx.MULTIPLE().getText();
            System.out.println("RegType File path: " + regType);
        }
    }

    @Override public void exitLogistic(ExprParser.LogisticContext ctx) {
        if (ctx.LOGISTIC() != null) {
            regType = ctx.LOGISTIC().getText();
            System.out.println("RegType File path: " + regType);
        }
    }

    @Override public void exitRidge(ExprParser.RidgeContext ctx) {
        if (ctx.RIDGE() != null) {
            regType = ctx.RIDGE().getText();
            System.out.println("RegType File path: " + regType);
        }
    }
    @Override public void exitSum(ExprParser.SumContext ctx) {
        if (ctx.TRUE() != null) {
            summary = true;
            String summaryText = ctx.TRUE().getText();
            System.out.println("Summary: " + summary);
        }
        if (ctx.FALSE() != null) {
            summary = false;
            String summaryText = ctx.FALSE().getText();
            System.out.println("Summary: " + summary);
        }
    }
    @Override public void exitListall(ExprParser.ListallContext ctx) {
        if (ctx.LISTALL() != null) {
            if (ctx.TRUE() != null) {
                Boolean listAll = true;
                String listAllText = ctx.TRUE().getText();
                System.out.println("List all: " + listAll);
            }
            if (ctx.FALSE() != null) {
                Boolean listAll = false;
                String listAllText = ctx.FALSE().getText();
                System.out.println("List all: " + listAll);
            }
        }
    }

    @Override public void exitSorttype(ExprParser.SorttypeContext ctx) {
        if (ctx.HIGHTOLOW() != null) {
            sortType = ctx.HIGHTOLOW().getText();
            System.out.println("Sort Type: High to Low");
        }
        if (ctx.ALPHA() != null) {
            sortType = ctx.ALPHA().getText();
            System.out.println("Sort Type: Alphabetical");
        }
        if (ctx.LOWTOHIGH() != null) {
            sortType = ctx.LOWTOHIGH().getText();
            System.out.println("Sort Type: Low to High");
        }
        if (ctx.BACKALPHA() != null){
             sortType = ctx.BACKALPHA().getText();
            System.out.println("Sort Type: Backward Alphabetical");
        }
    }

    //   private java.util.Vector<String> terms = new java.util.Vector<>();

}
public class EntireFile {
    public static void main(String[] args) throws Exception {
        CharStream input = CharStreams.fromFileName("inputs/test.txt");
        ExprLexer lexer = new ExprLexer(input);
        CommonTokenStream tokens = new CommonTokenStream(lexer);
        ExprParser parser = new ExprParser(tokens);
        LangListener lang = new LangListener();
        parser.addParseListener(lang);
        parser.prog();
        ExcelKeywordExtractor.main(lang.AssfilePath, lang.years, lang.depFilePath, lang.indFilePath, lang.regType, lang.summary, lang.graphType, lang.sortType);
        //System.out.println("Total number of rules: " + count.getCount());
        // System.out.println("Total number of List rules: " + count.getListCount());

    }
}
