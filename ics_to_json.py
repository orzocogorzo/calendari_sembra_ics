#! /usr/bin/env python3

import re
import json


def ics_reader(file_path):
    file = open(file_path, "r")

    line = file.readline()
    while line:
        yield line
        line = file.readline()

    file.close()
    yield False


def parse_ics_line(line):
    key = re.sub(r":.*$", "", line).strip()
    value = re.sub(r"^[^\:]*\:", "", line).strip()
    return [key, value]


def parse_ics_component(lines, is_event=False):
    data = dict()

    while line := next(lines):
        line = parse_ics_line(line)
        if line[0] == "BEGIN":
            data[line[1]] = data.get(line[1], [])
            data[line[1]].append(parse_ics_component(lines, line[1] == "VEVENT"))
        elif line[0] == "END":
            break
        else:
            data[line[0]] = line[1]
    
    if is_event:
        return format_event_data(data)

    return data

def format_event_data(data):
    taxonomies = dict()
    expanded = dict(TAXONOMIES=taxonomies)

    for key, value in data.items():
        if key == "DESCRIPTION":
            expanded["PLANT"] = re.match(r"<b>((\w| |')+)<\/b>", value).groups()[0].strip()
            taxonomies["EVENT"] = re.search(r"Esdeveniment\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["FAMILY"] = re.search(r"Familia\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["LIFECYCLE"] = re.search(r"Cicle de vida\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["SEEDING_DEPTH"] = re.search(r"Profunditat semba \(cm\)\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["GERMINATION_DAYS"] = re.search(r"Temps de germinació o brotació \(dies\)\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["HARVEST_TIME"] = re.search(r"Durada del cicle fins la recolecció\: ?([^<]+)<br\/>", value).groups()[0].strip()
            taxonomies["PLANTING_FRAME"] = re.search(r"Marc de plantació\: ?(.+)$", value).groups()[0].strip()
        elif key == "LOCATION":
            taxonomies["LOCATION"] = value.strip()
        
        expanded[key] = value

    return expanded

if __name__ == "__main__":
    from sys import argv

    reader = ics_reader(argv[1])
    data = parse_ics_component(reader)

    with open("ics.json", "w") as file:
        file.write(json.dumps(data, ensure_ascii=False))
