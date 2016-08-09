from pptx import Presentation

courseName = "What's New in Tableau 10"
courseSlug = "tableau-10-whats-new"

headerImg = 'profile.png'
authorName = 'Ben Sullins'
authorTitle = 'Data Geek'
authorFooter = '@bensullins  www.bensullins.com'


data = {

    "slides": [

        {"module_num":1, "slide_title":"What's New in Tableau 10", "slide_layout":0},
        {"module_num":1, "slide_title":"Course Overview", "slide_layout":2},
        {"module_num":1, "slide_title":"Relationship to Existing Courses", "slide_layout":2},
        {"module_num":1, "slide_title":"Our Story", "slide_layout":2},
        {"module_num":1, "slide_title":"Where to Find More", "slide_layout":2},
        {"module_num":2, "slide_title":"User Interface Updates", "slide_layout":0},
        {"module_num":2, "slide_title":"Reviewing the User Interface", "slide_layout":2},
        {"module_num":2, "slide_title":"Dashboarding for Mobile", "slide_layout":2},
        {"module_num":2, "slide_title":"Where to Find More", "slide_layout":2},
        {"module_num":3, "slide_title":"Preparing Your Data", "slide_layout":0},
        {"module_num":3, "slide_title":"Joining Across Databases", "slide_layout":2},
        {"module_num":3, "slide_title":"Connecting to Google Sheets", "slide_layout":2},
        {"module_num":3, "slide_title":"Connecting to QuickBooks", "slide_layout":2},
        {"module_num":3, "slide_title":"Combining Queries with Union", "slide_layout":2},
        {"module_num":3, "slide_title":"Where to Find More", "slide_layout":2},
        {"module_num":4, "slide_title":"Analyzing Your Data", "slide_layout":0},
        {"module_num":4, "slide_title":"Customizing Territories", "slide_layout":2},
        {"module_num":4, "slide_title":"Filtering Across Data Sources", "slide_layout":2},
        {"module_num":4, "slide_title":"Highlighting Data with Search", "slide_layout":2},
        {"module_num":4, "slide_title":"Sizing Marks", "slide_layout":2},
        {"module_num":4, "slide_title":"Calculating Radial Distance", "slide_layout":2},
        {"module_num":4, "slide_title":"Mapping New Areas", "slide_layout":2},
        {"module_num":5, "slide_title":"Advancing Your Analytics", "slide_layout":0},
        {"module_num":5, "slide_title":"Clustering Analysis", "slide_layout":2},
        {"module_num":5, "slide_title":"Updating Table Calculations", "slide_layout":2},
        {"module_num":5, "slide_title":"Specifying a Level of Detail in Calculations", "slide_layout":2},
        {"module_num":5, "slide_title":"Integrating Python Machine Learning", "slide_layout":2},
        {"module_num":5, "slide_title":"Where to Find More", "slide_layout":2},
        {"module_num":6, "slide_title":"Sharing Your Insights", "slide_layout":0},
        {"module_num":6, "slide_title":"Publishing Visualizations to Tableau Server", "slide_layout":2},
        {"module_num":6, "slide_title":"Subscribing Users to Your Visualizations", "slide_layout":2},
        {"module_num":6, "slide_title":"Viewing on Mobile", "slide_layout":2},
        {"module_num":6, "slide_title":"Analyzing Usage on Tableau Server", "slide_layout":2},


    ]

}



prs = Presentation('ps-template.pptx')

# Placeholders in slide 0
# 12, BODY (2)
# 14, BODY (2)
# 10, BODY (2)
# 15, PICTURE (18)
# 13, BODY (2)
# 0, TITLE (1)


prev_module = 0;
last_item = False

for i, slide in enumerate(data['slides']):

    print "working on item "+str(i)

    if i == len(data['slides'])-1:
        last_item = True

    cur_module = slide['module_num']

    # new module
    # save current preso, create new, add slide
    if last_item:

        # add slide
        section = prs.slides.add_slide(prs.slide_layouts[slide['slide_layout']])
        section_title = section.shapes.title
        section_title.text = slide['slide_title']

        prs.save(courseSlug + '-m' + str(cur_module) + '.pptx')
        print "created slides for module " + str(cur_module)

    elif cur_module == prev_module or prev_module == 0:

        # add slide
        section = prs.slides.add_slide(prs.slide_layouts[slide['slide_layout']])
        section_title = section.shapes.title
        section_title.text = slide['slide_title']

        # add author image
        if slide['slide_layout'] == 0:
            # print "attempting to add image 2"
            placeholder = section.placeholders[15]
            picture = placeholder.insert_picture(headerImg)
            section.placeholders[12].text = authorFooter
            section.placeholders[14].text = authorTitle
            section.placeholders[10].text = authorName
            # section.placeholders[13].text = '13' # module sub-title


    else:
        # save current preso
        prs.save(courseSlug + '-m' + str(prev_module) + '.pptx')

        # print msg
        print "created slides for module " + str(prev_module)

        # create new preso
        prs = Presentation('ps-template.pptx')

        # add slide
        section = prs.slides.add_slide(prs.slide_layouts[slide['slide_layout']])
        section_title = section.shapes.title
        section_title.text = slide['slide_title']

        # add author image
        if slide['slide_layout'] == 0:
            # print "attempting to add image 2"
            placeholder = section.placeholders[15]
            picture = placeholder.insert_picture(headerImg)
            section.placeholders[12].text = authorFooter
            section.placeholders[14].text = authorTitle
            section.placeholders[10].text = authorName
            # section.placeholders[13].text = '13' # module sub-title


    #set prev_module for iteration
    prev_module = cur_module