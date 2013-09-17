from docx import DocX, make_row
import re, random, string
import json

class DocXReplace(DocX):
    def __init__(self, input_filename, json_file = None, 
                       jsonstr = None, dic = None):
        super(DocXReplace, self).__init__(input_filename)
        if json_file is not None:
            f = open(json_file)
            self.replacements = json.loads(f.read())
            f.close()
        elif jsonstr is not None:
            self.replacements = json.loads(jsonstr)
        elif dic is not None:
            self.replacements = dic
        else:
            raise Exception("No data supplied to load_replacements")

        self.text_reps = self.replacements.get("text", {})
        self.table_reps = self.replacements.get("tables", {})
        self.image_reps = self.replacements.get("images", {})

    def replace_tags(self, line, replacements, specific_words = None):
        subs = re.split(r'(@[^@]*@)', line)
        res = ""
        count = 0
        for sub in subs:
            if self.is_key(sub):
                # if we've given a specific word list, and this isn't in it:
                if specific_words and sub[1:-1] not in specific_words:
                    res += sub # just append as-is and continue
                    continue
                try:
                    key = sub[1:-1]
                    self.log("replacing '%s' with '%s'" % (key, replacements[key]))
                    res += replacements[key].__str__()
                    count += 1
                except KeyError:
                    #if it's not in our lookup table, append as-is
                    if self.verbose:
                        print "Key '%s' not found in replacements!" % sub
                    res += sub
            else:
                res += sub
        return res, count

    def is_key(self, string):
        return len(string) > 2 and string[0] == string[-1] == '@'

    def replace_text(self, replacements = None, specific_words = None):
        ''' Finds and makes all of the replacements. '''
        if replacements is None:
            if self.text_reps is not None:
                replacements = self.text_reps
            else:
                raise Exception("No text replacements defined")
        document = self.get_document()
        count = 0
        for elem in document.iter():
            if elem.text:
                elem.text, c = self.replace_tags(elem.text, replacements, specific_words)
                count += c
        self.log("Made %d replacements" % count)

    def replace_image(self, imagename, new_image):
        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "graphic":
                rid = self.get_id(elem)
                picname = self.get_pic_name(elem)
                if picname == imagename and rid is not None:
                    self.set_image_relation(rid, new_image)
                    return
        # should probably throw exception if program flow reaches here

    def replace_images(self, replacements = None):
        if replacements is None:
            if self.image_reps is not None:
                replacements = self.image_reps
            else:
                raise Exception("No image replacements defined")
        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "graphic":
                picname = self.get_pic_name(elem)
                if picname and picname in replacements:
                    rid = self.get_id(elem)
                    if rid is not None:
                        self.log("Replacing %s with %s" % (picname, replacements[picname]))
                        self.set_image_relation(rid, replacements[picname])
                    else:
                        if self.verbose:
                            print "Relation id for image %s not present; can't replace" % picname

    def replace_tables(self, table_replacements = None):
        if table_replacements is None:
            if self.table_reps is not None:
                table_replacements = self.table_reps
            else:
                raise Exception("no table replacements dict defined")

        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "tbl":
                try:
                    nrows = self.get_num_rows(elem)
                    ncols = self.get_num_columns(elem)
                    for i in range(len(elem)):
                        if elem[i].tag.split("}")[-1] == "tr":
                            rowelem = elem[i]
                            col = self.find_subelem_list(rowelem, ["tc", "p", "r", "t"])

                            if col is None:
                                raise Exception("Could not find column")

                            tags = re.findall(r'@@([^@]+)@@', col.text)

                            if not tags:
                                # print "No @@ tag found to describe source in row %d" % i
                                continue
                            source = tags[0]
                            self.log("Found table tag %s, querying dictionary" % source)
                            if source not in table_replacements:
                                raise Exception("Error: couldn't find %s in replacements dict" % source)
                            settings = table_replacements[source][0]
                            font_size = settings.get("font_size", None)
                            # hack because fonts appear half-size for some reason
                            if font_size is not None:
                                font_size *= 2
                            font_face = settings.get("font_face", None)
                            # get the border settings
                            borders = settings.get("borders", [])

                            under_border = settings.get("under_border", False)
                            content = table_replacements[source][1]
                            
                            tbl_ncols = len(content[0])
                            if tbl_ncols != ncols:
                                raise Exception("Error: should have %d columns, but "
                                                "source has %d columns" % (ncols, tbl_ncols))
                            first = True
                            j = 0
                            for row in content:
                                if first:
                                    # if it's the first row we're appending, we want to
                                    # overwrite the row that was at i, a.k.a. the row
                                    # containing the @@tag@@.
                                    elem[i] = make_row(row, font_face=font_face, 
                                                            font_size=font_size.__str__(),
                                                            borders = borders)
                                    first = False
                                else:
                                    # otherwise we can just add to the end of the table
                                    elem.append(make_row(row, font_face=font_face,
                                                              font_size=font_size.__str__(),
                                                              borders = borders))
                                j += 1
                            self.log("Inserted %d rows into table %s" % (j, source))
                            break # only do it once for each table
                except Exception as e:
                    self.log(e + "\n" + "Error reading or constructing table element, no rows added")
                    return

    def replace_all(self, text_reps = None, table_reps = None, image_reps = None):
        self.log("replacing text...")
        self.replace_text(text_reps)
        self.log("replacing tables...")
        self.replace_tables(table_reps)
        self.log("replacing images...")
        self.replace_images(image_reps)
        self.log("done")