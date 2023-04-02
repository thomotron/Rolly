package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"golang.org/x/oauth2"
	"log"
	"math"
	"net/http"
	"os"
	"strconv"
)

// Logs a fatal error if err is not nil
func fatalErr(err error, message string) {
	if err != nil {
		if len(message) > 0 {
			log.Fatalf("%v: %v", message, err.Error())
		} else {
			log.Fatal(err.Error())
		}
	}
}

// Parse an A1 column string to a zero-based integer offset.
// For example, 'A' becomes 0, 'B' becomes 1, 'AE' becomes 30, etc.
func parseA1ColumnToInt(column string) (int, error) {
	if column == "" {
		return 0, errors.New("column string is empty")
	}

	output := 0
	runes := []rune(column)
	multiplier := 0
	for i := len(runes) - 1; i >= 0; i-- {
		// Enforce that all characters are alphabetical and uppercase
		if runes[i] < 'A' || runes[i] > 'Z' {
			return 0, errors.New("non-uppercase or non-alphabetical character in column string")
		}

		// Add to the output
		output += (int(math.Pow(float64(26), float64(multiplier))) * (int(runes[i]-'A') + 1))
		multiplier++
	}

	return output - 1, nil
}

// parseA1Notation parses the given A1 string and returns the x and y offsets of the range, followed by the width and height respectively.
func parseA1Notation(_range string) (int, int, int, int, error) {
	// Try parse the range using regex
	groups := a1RangePattern.FindStringSubmatch(_range)
	if groups == nil {
		return 0, 0, 0, 0, errors.New("input is not an A1 notation range")
	}
	startCol, startRow, endCol, endRow := groups[1], groups[2], groups[3], groups[4]

	// Parse start offsets
	x, err := parseA1ColumnToInt(startCol)
	if err != nil {
		return x, 0, 0, 0, fmt.Errorf("unable to parse range start column offset: %v", err)
	}

	y, err := strconv.Atoi(startRow)
	if err != nil {
		return x, y, 0, 0, fmt.Errorf("unable to parse range start row offset: %v", err)
	}
	y -= 1

	if endCol != "" && endRow != "" {
		// Parse the range end offsets as well
		xOffset, err := parseA1ColumnToInt(endCol)
		if err != nil {
			return x, y, 0, 0, fmt.Errorf("unable to parse range end column offset: %v", err)
		}

		yOffset, err := strconv.Atoi(endRow)
		if err != nil {
			return x, y, 0, 0, fmt.Errorf("unable to parse range end row offset: %v", err)
		} //yOffset -= 1

		return x, y, xOffset - x, yOffset - y, nil
	} else {
		// Assume single cell
		return x, y, 1, 1, nil
	}
}

// pluralise returns singular if n is 1, otherwise plural
func pluralise(singular string, plural string, n int) string {
	if n == 1 {
		return singular
	} else {
		return plural
	}
}

// Retrieves a token, saves it, then returns the generated client.
func getClient(tokenPath string, config *oauth2.Config) *http.Client {
	tok, err := tokenFromFile(tokenPath)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokenPath, tok)
	}
	return config.Client(context.Background(), tok)
}

// Requests a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

// hexToRGB converts the given hexadecimal colour value string to its component red, green, and blue values.
func hexToRGB(hex string) (int, int, int, error) {
	values, err := strconv.ParseUint(hex, 16, 32)

	if err != nil {
		return 0, 0, 0, err
	}

	return int(values >> 16), int((values >> 8) & 0xFF), int(values & 0xFF), nil
}
