import sys
from create_slides import SlideManager
from slide import Slide
import json
import logging

logging.basicConfig(filename='events.log',
                    filemode='w',
                    format='%(asctime)s: %(levelname)s - %(message)s',
                    datefmt='%yyyy-%mm-%dd %H:%M')


def import_file(filename):
    """Import presentation configuration from .json file\n
        """

    # create a list for the slide information
    slide_list = []

    # load data from json file
    try:
        logging.info("import_file: Start import datas from json file")
        with open(filename) as f:
            datas = json.load(f);

    # filename or path not exist
    except FileNotFoundError:
        print("File not found. Please check that the filename and the path is correct!")
        logging.error("import_file: File not found!")

    # catch unexpected error
    except Exception as e:
        logging.error("import_file: load data error: ", e)

    else:
        try:
            for data in datas['presentation']:
                # check that the block contain configuration dictionary
                if 'configuration' in data.keys():

                    # create slide class with configuration attribute
                    slide = Slide(slide_type=data['type'],
                                  title=data['title'],
                                  content=data['content'],
                                  configuration=data['configuration'])
                else:

                    # create slide class without configuration attribute
                    slide = Slide(slide_type=data['type'],
                                  title=data['title'],
                                  content=data['content'],
                                  configuration="")

                # add slide information to a list
                slide_list.append(slide)

        # catch unexpected error
        except Exception as e:
            logging.error("import_file: create slide class error: ", e)

        else:
            logging.info("import_file: Import successfully")
            return slide_list


def presentation(input_file, outputfile):
    """Generate presentation\n
        input_file:
            .json file which contains the slide information\n
        output_file:
            .ppt or .pptx file what is the presentation name"""

    logging.info("presentation: Start generate")

    # create a SlideManager object and add the .ppt file name
    slide_creator = SlideManager(outputfile);

    # get the slides information
    slides_list = import_file(input_file)

    # go through the slide information list
    for slide in slides_list:

        # get actual slide type
        try:
            stype = slide.slide_type
            if stype == "":
                raise ValueError("Slide type is not correct")

        # don't add the slide type
        except ValueError:
            logging.error("presentation: the file not contained the slide type.")

        # handling unexpected error
        except Exception as e:
            logging.error("presentation: slide type error: ", e)

        else:
            try:
                # create the slide
                if stype == 'title':
                    slide_creator.title_slide(slide)
                elif stype == 'text':
                    slide_creator.text_slide(slide)
                elif stype == 'list':
                    slide_creator.list_slide(slide)
                elif stype == 'picture':
                    slide_creator.image_slide(slide)
                elif stype == 'plot':
                    slide_creator.plot_slide(slide)
                else:
                    raise TypeError

            # type is not correct
            except TypeError:
                logging.error("Presentation: this type of slide is not exist")
                print(stype, ": this type of slide is not exist")

            # handling unexpected error
            except Exception as e:
                logging.error("presentation: slide create error: ", e)

            else:
                logging.info("presentation: Generate successfully")


if __name__ == "__main__":
    try:
        # get the filename from the arguments
        file = str(sys.argv[1])

        # check that the file is .json
        if ".json" not in file:
            raise NameError

    # missing argument
    except IndexError:
        print("Please add the .json file in arguments. Correct run is 'python create_presentation.py sample.json'")
        logging.error("main: missing filename argument")

    # wrong extension
    except NameError:
        print("The file extension is wrong! Example: 'sample.json'")
        logging.error("main: The file is not .json")

    # create presentation
    else:
        presentation(file, 'sample.pptx')
