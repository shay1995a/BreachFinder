import requests
import sys
import xlsxwriter
import argparse


def parse_args():
    # parsing blat
    parser = argparse.ArgumentParser(prog='breachFinder',
                                     description='A tool to work with hunter.io and excel')
    parser.add_argument('-d', '--domain',
                        action='store',
                        dest='domain',
                        help='give a domain name to scan example: kaki.com (no need with http or www)',
                        type=str,
                        required=True)
    parser.add_argument('-hapi', '--hunterapi',
                        action='store',
                        dest='hunterapi',
                        help='give the HUNTER.IO API-key please',
                        type=str,
                        required=True)
    parser.add_argument('-o', '--output-file',
                        action='store',
                        dest='output',
                        help='name the file for the excel',
                        type=str,
                        required=True)
    return parser.parse_args()


def export_results(emails, file_name):
    # export results to xlsx file
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0
    count = 0
    try:
        print("exporting results to an excel file")
        workbook = xlsxwriter.Workbook(file_name + '.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write(row, col, "email")
        row += 1

        # go over the data and write it to a new row
        for email in emails:
            col = 0
            worksheet.write(row, col, email)
            row += 1
            count += 1
        # close the excel
        workbook.close()

    except Exception as excp:
        print("Something went wrong with the export Blat\n"+str(excp))


def edit_response(data):
    # filter the response from the api
    emails = []
    try:
        for email in data['data']['emails']:
            print("\n[*]Email found: " + str(email['value']))
            emails.append(str(email['value']))
    except Exception:
        print("Could not find any info about that domain or Change the API key")
        emails = '-'
    return emails


def send_request(url):
    response = None
    try:
        response = requests.get(url, timeout=5, allow_redirects=True)
    except Exception as excp:
        print(excp)
    return response.json()


def main(args):
    target = args.domain
    api = args.hunterapi
    file_name = args.output
    emails = []
    limit = 100
    # this is the maximum you can get
    try:
        url = "https://api.hunter.io/v2/domain-search?domain="+target+"&api_key="+api+"&limit="+str(limit)
#if you want more then the first 100 emails add the offset parameter to the request example :< &offset=100 > will get you the next 100
        # sent request
        response = send_request(url)
        # edit response
        emails = edit_response(response)
        # Export results
        if emails != "-":
            export_results(emails, file_name)
    except Exception as exception:
        print("Error in main function" + str(exception))


if __name__ == "__main__":
    main(parse_args())
