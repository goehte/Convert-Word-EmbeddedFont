#include <stdio.h> 
#include <stdlib.h> 
#include <stdint.h>
#include <stdbool.h>
#include <string.h>

// #define DEBUG

// Programm Name: docx_getembfnt
// WordML - File Embedded Word Front: re-obfuscate font tool 

// This C programm is generated based in the following forum thread: 
// https://forum.softmaker.de/viewtopic.php?f=275&t=25741&p=126388#p123933

// [ECMA-376-1] 
// Documentantion for Office Open XML File Format ECMA-376-1:2016:
// https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
// Part 1: Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
// Section 17.8.1, Page 670

// GUID handling - obfuscation
// Example: 001B70DC-AA60-4AD5-90EC-18A0948E1EAE
// To:      AE1E8E94-A018-EC90-D54A-60AADC701B00

// Build under Linux:
// $ gcc main.c -o docx_getembfnt

// TOOL USAGE
//
// 1. Open *.docx file as ZIP, most simple just rename it to *.zip
//
// 2. Unzip obfuscated fonts from the following folder: /word/fonts/ 
//    Example file name of a obfuscated TTF font: font1.odttf
//
// 3. Open /word/fontTable.xml in a editor or XML viewer
//
// 4. Search for related "fontKey" GUID in <w:embedRegular> . e.g. "font1.ottf" == "rId1"  
//    Example: <w:embedRegular w:fontKey="{7B19B49C-2336-4F82-AAD2-5D2BAE389560}" r:id="rId1"/>
// 
// 5. Run: 
//    Linux: $ docx_getembfnt <inputfile> <outputfile> <GUID>
//    Windows $ docx_getembfnt32.exe <inputfile> <outputfile> <GUID>
//    Examples usage (Linux):
//     $ ./docx_getembfnt ./examples/font1.odttf ./examples/font1.ttf 7B19B49C-2336-4F82-AAD2-5D2BAE389560
//     $ ./docx_getembfnt ./examples/font2.odttf ./examples/font2.ttf 373EDBC8-20C2-49CE-A04A-3A982C578A05
//     $ ./docx_getembfnt ./examples/font3.odttf ./examples/font3.ttf 3E0C3015-B393-4155-9571-A242AB103A61
//     $ ./docx_getembfnt ./examples/font4.odttf ./examples/font4.ttf 46483AA3-95DD-4E29-834A-6A30BA96940D
//
// 6. Output file you should able to insall as new font to your system.

// Note: Sure the tool works also "vi­ce ver­sa". So you can make from a font file a obfuscated file.

// This source-code is considered to be public-domain (freeware) without any liability and warranty.
// -goehte80, Erlangen, 12-Apr-2021


// Open file function
static FILE *open_file ( char *file, char *mode )
{
  FILE *fp = fopen ( file, mode );

  if ( fp == NULL ) {
    perror ( "Unable to open file" );
    exit ( EXIT_FAILURE );
  }

  return fp;
}

// Read full hex file function
size_t read_hex_file(FILE *file, unsigned char *dest) {
  size_t count = 0;
  int n;
  if (dest == NULL) {
    unsigned char OneByte;
    while ((n = fgetc(file)) != EOF) {
    //while ((n = fscanf(inf, "%c", &OneByte)) == 1 ) {
      count++;
    }
  }
  else {
    while ((n = fgetc(file)) != EOF) {
    //while ((n = fscanf(inf, "%c", dest)) == 1 ) {
      dest[count] = n;
      count++;
    }
  }
  if (n != EOF) {
    ;  // handle syntax error
  }
  return count;
}

// Write full hex file function
 size_t write_font_hex_file(FILE *file, unsigned char guid_hex[16], unsigned char *dest, size_t count) {
  size_t n;
  if (dest == NULL) {
    ; // do nothing
  }
  else if (count == 0) {
    ; // do nothing
  }
  else {
    printf("\n");
    // HEX write: first 32 bytes of the binary: 0-15
    for (n = 0; n <= 15; n++) {
      fputc((guid_hex[n]^dest[n]), file);
      printf("Write pos %02lu: GUID: 0x%02X HEX: 0x%02X XOR: 0x%02X\n", n, guid_hex[n], dest[n], guid_hex[n]^dest[n]);
    }
    printf("\n");
    // HEX write: secoundt 32 bytes of the binary: 16-31
    for (n = 16; n <= 31; n++) {
      fputc((guid_hex[n-16]^dest[n]), file);
      printf("Write pos %02lu: GUID: 0x%02X HEX: 0x%02X XOR: 0x%02X\n", n, guid_hex[n-16], dest[n], guid_hex[n-16]^dest[n]);
    }
    // HEX write: remaning file
    for (n = 32; n <= (count-1); n++) {
      fputc(dest[n], file);
    }
    printf("\n");
  }
  return n;
}

// String slice function
// Source: https://stackoverflow.com/questions/20342559/how-to-cut-part-of-a-string-in-c
/**
 * Extracts a selection of string and return a new string or NULL.
 * It supports both negative and positive indexes.
 */
const char * str_slice(char str[], int slice_from, int slice_to)
{
    // if a string is empty, returns nothing
    if (str[0] == '\0')
        return NULL;

    char *buffer;
    size_t str_len, buffer_len;

    // for negative indexes "slice_from" must be less "slice_to"
    if (slice_to <= 0 && slice_from < slice_to) {
        str_len = strlen(str);

        // if "slice_to" goes beyond permissible limits
        if (abs(slice_to) > str_len - 1)
            return NULL;

        // if "slice_from" goes beyond permissible limits
        if (abs(slice_from) > str_len)
            slice_from = (-1) * str_len;

        buffer_len = slice_to - slice_from;
        str += (str_len + slice_from);

    // for positive indexes "slice_from" must be more "slice_to"
    } else if (slice_from >= 0 && slice_to > slice_from) {
        str_len = strlen(str);

        // if "slice_from" goes beyond permissible limits
        if (slice_from > str_len - 1)
            return NULL;

        buffer_len = slice_to - slice_from;
        str += slice_from;

    // otherwise, returns NULL
    } else
        return NULL;

    buffer = calloc(buffer_len, sizeof(char));
    strncpy(buffer, str, buffer_len);
    return buffer;
}


// GUID String to Hex (with reversing the order)
// see [ECMA-376-1] 17.8.1 obfuscation algorithm
void guid_to_reverse_hex (char guid_str[36], unsigned char guid_hex[16])
{

    // String to Hex - Example:
    /*
    char test_hex[] = "6A";                          // here is the example hex string
    unsigned char test_num = (unsigned)strtol(test_hex, NULL, 16);       // number base 16
    printf("%c\n", test_num);                        // print it as a char
    printf("%d\n", test_num);                        // print it as decimal
    printf("%X\n", test_num);                        // print it back as hex
    */

    // String to reverse order Hex conversation with "str_slice" function
    // Note: sure there is may a better, more efficent solution insted of using "str_slice" function- Feel free to improve.

    guid_hex[0] = (unsigned)strtol((str_slice(guid_str, 34, 36)), NULL, 16);
    guid_hex[1] = (unsigned)strtol((str_slice(guid_str, 32, 34)), NULL, 16);
    guid_hex[2] = (unsigned)strtol((str_slice(guid_str, 30, 32)), NULL, 16);
    guid_hex[3] = (unsigned)strtol((str_slice(guid_str, 28, 30)), NULL, 16);
    guid_hex[4] = (unsigned)strtol((str_slice(guid_str, 26, 28)), NULL, 16);
    guid_hex[5] = (unsigned)strtol((str_slice(guid_str, 24, 26)), NULL, 16);
    guid_hex[6] = (unsigned)strtol((str_slice(guid_str, 21, 23)), NULL, 16);
    guid_hex[7] = (unsigned)strtol((str_slice(guid_str, 19, 21)), NULL, 16);
    guid_hex[8] = (unsigned)strtol((str_slice(guid_str, 16, 18)), NULL, 16);
    guid_hex[9] = (unsigned)strtol((str_slice(guid_str, 14, 16)), NULL, 16);
    guid_hex[10] = (unsigned)strtol((str_slice(guid_str, 11, 13)), NULL, 16);
    guid_hex[11] = (unsigned)strtol((str_slice(guid_str, 9, 11)), NULL, 16);
    guid_hex[12] = (unsigned)strtol((str_slice(guid_str, 6, 8)), NULL, 16);
    guid_hex[13] = (unsigned)strtol((str_slice(guid_str, 4, 6)), NULL, 16);
    guid_hex[14] = (unsigned)strtol((str_slice(guid_str, 2, 4)), NULL, 16);
    guid_hex[15] = (unsigned)strtol((str_slice(guid_str, 0, 2)), NULL, 16);
}


// MAIN
int main(int argc, char *argv[])
{

 // Variable definitions:
  char *filename_in, *filename_out; // File name to open
  char version[32] = "V0.1 April-2021";

  unsigned char guid_hex[16];
  char guid_str[37];

  // File handle:
	FILE *fh_hex_file;

  system("clear"); // Clear Screen
  printf("\nProgramm Name: docx_getembfnt [%s]\n", version);
  printf("WordML - File Embedded Word Front: re-obfuscate font tool\n");
  printf("=========================================================\n\n");
  
#ifdef DEBUG

	if ( argc == 1 ) {
    // if no filename argument is given
  	filename_in = "./debug/word_font3.odttf";
    filename_out = "./debug/word_font3.ttf";
    strcpy(guid_str, "0A92BF73-9C23-4422-BBAF-6E7D93B78DE");
  }

#else
  // Command argument handling for filename
  if ( argc != 4 ) {
  //If the user fails to give us two arguments yell at him.	
		fprintf ( stderr, "Usage: %s <inputfile> <outputfile> <GUID>\n", argv[0] );
		exit ( EXIT_FAILURE );
	} else if (argc == 4) {
  	filename_in = argv[1];
    filename_out = argv[2];
    strcpy(guid_str, argv[3]);
  } 
#endif

  // Print input informations
  printf("Input file name: %s\n", filename_in);
  printf("Output file name: %s\n", filename_out);
  printf("GUID: %s\n", guid_str);

  // GUID from frontTable.xml
  printf("fontKey GUID from frontTable.xml: %s\n", guid_str);

  // Reverse GUID string and format as HEX
  guid_to_reverse_hex(guid_str, guid_hex); 

  // Open file for reading
	fh_hex_file = open_file (filename_in, "rb" );
  
  // Read size of input file
  size_t hex_n = read_hex_file(fh_hex_file, NULL); // get size of file
  printf("Used input file name: %s\n", filename_in);
  printf("Size of input file: %u\n",(unsigned) hex_n);

  unsigned char *hex = malloc(hex_n); // reverse memory

  // Read data from obfuscated input font file
  rewind(fh_hex_file); // put file point to beginning of file
  read_hex_file(fh_hex_file, hex); // read file

  // Close file handle
  fclose(fh_hex_file);

  // Open file for writing
  fh_hex_file = open_file (filename_out, "wb" );

  // Write output front file
  size_t fhwcnt = write_font_hex_file(fh_hex_file, guid_hex, hex, hex_n); 

  // Close file handle
  fclose(fh_hex_file); 

  if (hex_n == fhwcnt) printf("Output file: %s sucessfully written\n", filename_out);

  free(hex); // free memory

  printf("Done\n");
  return 0;
}