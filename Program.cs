using System;
using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;


//This is going to be the main driver for the app!
namespace SlideSanitizer3000
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            pptxDoco Document = null;
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                // Values are available here
                if (options.inputFile!=null)
                {
                    Document = new pptxDoco(options.inputFile);
                }
                if (options.DeleteSlides!=null) DeleteSlidesWithTitle(options,false,ref Document);
                if (options.DeleteNotes) DeleteNotes(ref Document);
                if (options.Template!=null) ChangeTemplate(options, ref Document);
                if (options.DeleteContaining != null) DeleteSlidesWithTitle(options, true,ref Document);             
            }
            if ( Document!=null)
                Document.closefile();
        }
        static int DeleteSlidesWithTitle(Options options, bool containing, ref pptxDoco docu)
        {
            if (docu!=null)
            {
                docu.CountSlides();
                Console.WriteLine("Number of slides in this powerpoint: {0}" , docu.getCount());
            }
            List<int> indicies = new List<int>();
            int count = 0;
            //Get all the titles in the presentation
            if (containing)
            {
                Console.WriteLine("Input file is: {0} Deleting all slides with: {1} in the title",options.inputFile,options.DeleteContaining);
                foreach (string s in docu.getSlideTitles())
                {
                    if (s.Contains(options.DeleteContaining))
                    {
                        indicies.Add(count);
                    }
                    count++;
                }
            }
            else
            {
                Console.WriteLine("Input file is: {0} Deleting all slides with this exact string in the title: {1}",options.inputFile,options.DeleteSlides);
                foreach (string s in docu.getSlideTitles())
                {
                    if (options.DeleteSlides == s)
                    {
                        indicies.Add(count);
                    }
                    count++;
                }
            }
            //Delete all slides in the array impliment a stack to get this done
            indicies.Reverse();
            //use a for each to go through the list and delete from the end
            foreach (int a in indicies)
            {
                Console.WriteLine("Deleting slide at index: {0}", a);
                docu.deleteSlide(a);
            }
            //Check to make sure that the index is in the correct range (greater than 0 less than number of slides-1)
            return 0;
        }

        static int DeleteNotes (ref pptxDoco docu)
        {
            Console.WriteLine("Deletes all  the notes from slides!!!!!");
            try
            {
                docu.deleteNotes();
            }
            catch
            {
                Console.WriteLine("Something went wrong here!");
                return 1;
            }           
            return 0;
        }

        static int ChangeTemplate (Options option, ref pptxDoco docu)
        {
            Console.WriteLine("Input file is: {0} and the template to change to is: {1}", option.inputFile, option.Template);
            //Check to make sure document is not null, and make sure that the template is the correct type
            try
            {
                docu.changeTemplate(option.Template);
            }
            catch
            {
                Console.WriteLine("Something went wrong here!");
                return 2;
            }
            
            return 0;
        }

        static int tester (ref pptxDoco docu)
        {
            string szOutput;
            for (int i=0; i < docu.getCount(); i++)
            {
                docu.getSlideIDAndText(out szOutput, i);
                Console.WriteLine("The slide at index: {0} has this text: {1}", i, szOutput);
            }
            return 0;
        }
    }
    
    class Options
    {
        [Option('d', "delete-Slides",
            HelpText = "Option will delete slides with the supplied string only in the title")]
        public string DeleteSlides { get; set;}

        [Option('c',"Delete-Containing",
            HelpText = "option will delete slides with the specified string in the title")]
        public string DeleteContaining { get; set;}

        [Option('n',"delete-Notes",
            HelpText = "Using this option will delete all notes from the presentation")]
        public bool DeleteNotes { get; set;}

        [Option('t',"template",
            HelpText = "Option will change the template to one specified")]
        public string Template { get; set;}

        [Option('i', "Input-File", Required = true,
            HelpText = "Specifies the powerpoint file you want to run this against")]
        public string inputFile { get; set;}

        /*
        [Option('x', "tester-string",
            HelpText = "This is for testing methods!")]
        public bool tester { get; set; }
        */
        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
