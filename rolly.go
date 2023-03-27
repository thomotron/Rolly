package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/BurntSushi/toml"
	"github.com/bwmarrin/discordgo"
	set "github.com/deckarep/golang-set/v2"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"math"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
)

type Config struct {
	Google struct {
		CredentialsPath  string
		TokenPath        string
		RedirectURL      string
		SheetID          string
		SheetRangesSlice []string `toml:"SheetRanges"`
		sheetRanges      set.Set[string]
	}

	Discord struct {
		ApplicationID     string
		BotToken          string
		BotOwners         []string
		BotServer         string
		RollCallChannelID string
		ReactionColours   map[string]ColourPriority
	}
}

type ColourPriority struct {
	Colour   string
	Priority int
}

type NameColourUpdate struct {
	Name   string
	Colour string
}

type SlashCommand struct {
	Command *discordgo.ApplicationCommand
	Handler func(session *discordgo.Session, i *discordgo.InteractionCreate)
}

// This should be able to match all valid A1 ranges except for whole columns (A:A, B:C, etc.)
var a1RangePattern = regexp.MustCompile("(?:(?:'[\\w\\s]+'|[\\w\\s]+)!)?([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?")

var config *Config

// Define various chat commands
var commands = []*SlashCommand{
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "help",
			Description: "Prints this help text",
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "help") {
				return
			}

			err := session.InteractionRespond(i.Interaction, &discordgo.InteractionResponse{
				Type: discordgo.InteractionResponseChannelMessageWithSource,
				Data: &discordgo.InteractionResponseData{
					Content: "Hey, I'm Rolly! I create roll call messages in a channel you choose, and update the colours of users that react to them in a Google Sheets spreadsheet.\n" +
						"You can try sending a roll call with `/create`, or use one of the other commands to configure how I work.",
				},
			})
			if err != nil {
				fmt.Fprintf(os.Stderr, "Failed responding to interaction: %v\n", err)
			}
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "create",
			Description: "Creates a new roll call message",
			Options: []*discordgo.ApplicationCommandOption{
				{
					Type:        discordgo.ApplicationCommandOptionChannel,
					Name:        "channel",
					Description: "Channel to create the roll call in",
					Required:    false,
				},
				{
					Type:        discordgo.ApplicationCommandOptionString,
					Name:        "message",
					Description: "Message to use in the roll call",
					Required:    false,
				},
			},
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "create") {
				return
			}

			// Parse the options we've been given
			var channel *discordgo.Channel
			var rollCallMessage string
			for _, option := range i.Interaction.ApplicationCommandData().Options {
				switch option.Name {
				case "channel":
					channel = option.ChannelValue(nil)
					break
				case "message":
					rollCallMessage = option.StringValue()
					break
				}
			}

			// Try get the default channel, if there is one
			if rollCallMessage == "" {
				// Set the default message
				rollCallMessage = "Roll call!"
			}

			if channel == nil {
				var err error
				channelID := config.Discord.RollCallChannelID
				if channelID == "" {
					channelID = i.Interaction.ChannelID
				}
				channel, err = session.Channel(channelID)

				if err != nil {
					messageContent := fmt.Sprintf("I couldn't get the channel to send the roll call in. This is the message you gave me: `%s`", rollCallMessage)
					discordInteractionRespond(session, i.Interaction, &messageContent)
					return
				}
			}

			// Respond to the command
			channelName := fmt.Sprintf("<#%s>", channel.ID)
			messageContent := fmt.Sprintf("Creating a new roll call in %s...", channelName)
			discordInteractionRespond(session, i.Interaction, &messageContent)

			// Defer a message update for when we're done
			messageUpdateContent := fmt.Sprintf("Created a new roll call in %s with the following message: `%s`", channelName, rollCallMessage)
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Send the roll call message in the predefined channel
			message, err := session.ChannelMessageSend(channel.ID, rollCallMessage)
			if err != nil {
				messageUpdateContent = fmt.Sprintf("I couldn't make the roll call message. This is the message you gave me: `%s`", rollCallMessage)
				return
			}

			// Add reactions to the message
			for emoji := range (*config).Discord.ReactionColours {
				err = session.MessageReactionAdd(channel.ID, message.ID, emoji)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Failed adding %v emoji to roll call message: %v\n", emoji, err)
					messageUpdateContent = fmt.Sprintf("Created a new roll call in %s with the following message: `%s`.\nI couldn't add one or more emoji to the roll call message though.", rollCallMessage, channelName)
				}
			}
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "sheet",
			Description: "Gets the Google Sheets spreadsheet URL that is used for updates",
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "sheet") {
				return
			}

			// Respond to the command
			messageContent := "Getting sheet ID..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			// Defer a message update for when we're done
			messageUpdateContent := "I couldn't get the sheet ID"
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			if config.Google.SheetID == "" {
				messageUpdateContent = "I don't have a sheet ID set, so I won't be able to do any name updates. You can give me one with `/setsheet`."
			} else {
				messageUpdateContent = fmt.Sprintf("This is the sheet I'll use when updating name colours: https://docs.google.com/spreadsheets/d/%s/", config.Google.SheetID)
			}
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "setsheet",
			Description: "Sets the Google Sheets spreadsheet ID to update",
			Options: []*discordgo.ApplicationCommandOption{
				{
					Type:        discordgo.ApplicationCommandOptionString,
					Name:        "sheet-id",
					Description: "Google Sheets spreadsheet ID",
					Required:    true,
				},
			},
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "setsheet") {
				return
			}

			// Respond to the command
			messageContent := "Setting sheet ID..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			// Defer a message update for when we're done
			messageUpdateContent := "I couldn't update the sheet ID"
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Get the sheet ID
			sheetId := i.ApplicationCommandData().Options[0].StringValue()

			// Set the sheet ID
			config.Google.SheetID = sheetId
			messageUpdateContent = fmt.Sprintf("Set the sheet ID to `%s`", sheetId)
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "ranges",
			Description: "Shows the current allowed ranges in the spreadsheet",
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "ranges") {
				return
			}

			messageContent := "Getting ranges..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			messageUpdateContent := "I couldn't get the sheet ranges."
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Respond with the ranges
			ranges := config.Google.sheetRanges.ToSlice()
			if len(ranges) == 0 {
				messageUpdateContent = fmt.Sprintf("I don't have any ranges to look for matches in. You can add some with `/addrange` or `/setranges`.")
			} else {
				messageUpdateContent = fmt.Sprintf("Here %s the current %s I'll look for matches in: `%s`", pluralise("is", "are", len(ranges)), pluralise("range", "ranges", len(ranges)), strings.Join(ranges, "`, `"))
			}
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "addrange",
			Description: "Adds an allowed range for the spreadsheet",
			Options: []*discordgo.ApplicationCommandOption{
				{
					Type:        discordgo.ApplicationCommandOptionString,
					Name:        "range",
					Description: "Cell range represented in A1 notation (e.g. `E8`, `C2:D17`, `'My Other Sheet'!AE2:AF357`, etc.)",
					Required:    true,
				},
			},
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "addrange") {
				return
			}

			// Respond to the command
			messageContent := "Adding range..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			// Update the command
			messageUpdateContent := "Done!"
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Get and add the range
			options := i.Interaction.ApplicationCommandData().Options
			var value string
			if len(options) > 0 {
				value = options[0].StringValue()

				// Try parse the range before adding it
				_, _, _, _, err := parseA1Notation(value)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Invalid A1 notation range from addrange: %v", err)
					messageUpdateContent = fmt.Sprintf("`%s` doesn't look like a valid range in A1 notation. Check these examples from Google to see what I mean: https://developers.google.com/sheets/api/guides/concepts#expandable-1", value)
					return
				}

				config.Google.sheetRanges.Add(value)
			}

			messageUpdateContent = fmt.Sprintf("Added `%s` to the range list. %s that I'll look for matches in now: `%s`", value, pluralise("This is the one", "These are the ones", len(config.Google.sheetRanges.ToSlice())), strings.Join(config.Google.sheetRanges.ToSlice(), "`, `"))
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "setranges",
			Description: "Sets the ranges to update in the spreadsheet",
			Options: []*discordgo.ApplicationCommandOption{
				{
					Type:        discordgo.ApplicationCommandOptionString,
					Name:        "ranges",
					Description: "One or more ranges in A1 notation separated by a space (e.g. `E8 F2 G12`, `C2:D7 E2:G9 A1:D1`, etc.)",
					Required:    true,
				},
			},
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "setranges") {
				return
			}

			messageContent := "Setting ranges..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			messageUpdateContent := "I couldn't set the ranges."
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Unpack the ranges and try add them on
			options := i.Interaction.ApplicationCommandData().Options
			newRanges := set.NewSet[string]()
			if len(options) > 0 {
				matches := a1RangePattern.FindAllStringSubmatch(options[0].StringValue(), -1)
				for matchIndex, match := range matches {
					value := strings.TrimSpace(match[0])

					// Validate the range
					_, _, _, _, err := parseA1Notation(value)
					if err != nil {
						messageUpdateContent = fmt.Sprintf("I got %d %s in but `%s` doesn't look like a valid range in A1 notation", matchIndex, pluralise("match", "matches", matchIndex+1), match)
						return
					}

					newRanges.Add(value)
				}
			} else {
				return
			}

			// Replace the ranges
			newRangesSlice := newRanges.ToSlice()
			config.Google.sheetRanges.Clear()
			config.Google.sheetRanges.Append(newRangesSlice...)

			switch len(newRangesSlice) {
			case 0:
				messageUpdateContent = fmt.Sprintf("I couldn't find any valid A1 notation ranges in the list you gave me: `%s`")
				return
			case 1:
				messageUpdateContent = fmt.Sprintf("Replaced the range list with `%s`.", newRangesSlice[0])
				break
			default:
				messageUpdateContent = fmt.Sprintf("Replaced the range list with the following: `%s`", strings.Join(newRangesSlice, "`, `"))
				break
			}
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "channel",
			Description: "Shows the default channel to create roll calls in",
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "channel") {
				return
			}

			messageContent := "Getting roll call channel..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			messageUpdateContent := "I couldn't get the default roll call channel."
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			var channelName string
			if config.Discord.RollCallChannelID == "" {
				channelName = "whichever channel a roll call is created in"
			} else {
				channelName = fmt.Sprintf("<#%s>", config.Discord.RollCallChannelID)
			}
			messageUpdateContent = fmt.Sprintf("Currently roll calls will be sent in %s.\nUse `/setchannel` and mention a channel to set it, or just use `/setchannel` without any parameters to use whichever channel a roll call is created in.", channelName)
		},
	},
	{
		Command: &discordgo.ApplicationCommand{
			Name:        "setchannel",
			Description: "Sets the channel to create roll calls in",
			Options: []*discordgo.ApplicationCommandOption{
				{
					Type:        discordgo.ApplicationCommandOptionChannel,
					Name:        "channel",
					Description: "Channel to create roll calls in",
					ChannelTypes: []discordgo.ChannelType{
						discordgo.ChannelTypeGuildText,
					},
					Required: false,
				},
			},
		},
		Handler: func(session *discordgo.Session, i *discordgo.InteractionCreate) {
			if !assertApplicationCommand(i, "setchannel") {
				return
			}

			messageContent := "Setting roll call channel..."
			discordInteractionRespond(session, i.Interaction, &messageContent)

			messageUpdateContent := "I couldn't set the roll call channel."
			defer discordInteractionUpdate(session, i.Interaction, &messageUpdateContent)

			// Set the provided channel ID
			options := i.Interaction.ApplicationCommandData().Options
			if len(options) > 0 {
				channelID := options[0].ChannelValue(nil).ID
				config.Discord.RollCallChannelID = channelID
				messageUpdateContent = fmt.Sprintf("Set the roll call channel to <#%s>.", channelID)
			} else {
				config.Discord.RollCallChannelID = ""
				messageUpdateContent = "Set the roll call channel to be whichever channel a roll call is created in."
			}
		},
	},
}

// Loads config in from the given path
func loadConfig(path string) (*Config, error) {
	var config Config

	// Set defaults
	if len(config.Discord.ReactionColours) == 0 {
		config.Discord.ReactionColours = map[string]ColourPriority{
			"✅": {Colour: "00ff00", Priority: 1},
			"❔": {Colour: "ffff00", Priority: 2},
			"❌": {Colour: "ff0000", Priority: 3},
		}
	}
	if config.Google.CredentialsPath == "" {
		config.Google.CredentialsPath = "credentials.json"
	}
	if config.Google.TokenPath == "" {
		config.Google.TokenPath = "token.json"
	}

	// Try read in from the given path
	_, err := toml.DecodeFile(path, &config)
	if err != nil {
		return &config, err
	}

	// Check for required values
	if config.Google.RedirectURL == "" {
		return &config, errors.New("missing Google redirect URL")
	}
	if config.Discord.BotToken == "" {
		return &config, errors.New("missing Discord bot token")
	}
	if config.Discord.BotServer == "" {
		return &config, errors.New("missing Discord server ID")
	}

	// Parse computed values
	config.Google.sheetRanges = set.NewSet[string]()
	for _, sheetRange := range config.Google.SheetRangesSlice {
		config.Google.sheetRanges.Add(sheetRange)
	}

	return &config, nil
}

// Saves config to the given path
func saveConfig(config *Config, path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to open config file for writing: %v", err)
	}

	e := toml.NewEncoder(file)
	err = e.Encode(config)
	if err != nil {
		return fmt.Errorf("failed writing config to \"%s\": %v", path, err)
	}

	err = file.Sync()
	if err != nil {
		return fmt.Errorf("failed syncing file \"%s\": %v", path, err)
	}
	err = file.Close()
	if err != nil {
		return fmt.Errorf("failed closing file \"%s\": %v", path, err)
	}

	return nil
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

func main() {
	// Try loading config
	configPath := "config.toml"
	_config, loadErr := loadConfig(configPath)
	if loadErr != nil {
		if errors.Is(loadErr, os.ErrNotExist) {
			fmt.Printf("Config file \"%s\" doesn't exist, creating a new one.\nPlease fill it out accordingly and run the program again.\n", configPath)
			saveErr := saveConfig(_config, configPath)
			if saveErr != nil {
				fatalErr(saveErr, fmt.Sprintf("Failed saving config to \"%s\"", configPath))
			}
			os.Exit(1)
		} else {
			fatalErr(loadErr, fmt.Sprintf("Failed loading config from \"%s\"", configPath))
		}
	}
	config = _config

	// Save config on close
	defer func() {
		saveErr := saveConfig(config, configPath)
		if saveErr != nil {
			fatalErr(saveErr, fmt.Sprintf("Failed saving config to \"%s\"", configPath))
		}
	}()

	// Initialise API services
	sheetsService, err := initSheets(config.Google.CredentialsPath, config.Google.TokenPath)
	fatalErr(err, "Failed to initialise sheets service")
	discordService, err := initDiscord(config.Discord.BotToken)
	fatalErr(err, "Failed to initialise Discord session")

	// Create a name update queue to be shared by the Discord event handlers and the Sheets update loop
	nameColourUpdates := make(chan NameColourUpdate, 10)

	// Register Discord event handlers
	err = registerDiscordEvents(discordService, config, nameColourUpdates)
	fatalErr(err, "Failed to register Discord event handlers")

	fmt.Println("Ready!")

	// Run the sheet event loop
	sheetTicker := time.NewTicker(3_000_000_000)
	for {
		select {
		case <-sheetTicker.C:
			processQueue(nameColourUpdates, sheetsService, config.Google.SheetID, config.Google.sheetRanges.ToSlice())
		}
	}
}

// processQueue will consume all available items in the given queue and push the changes using the given sheetsService instance.
func processQueue(queue <-chan NameColourUpdate, sheetsService *sheets.Service, sheetID string, sheetRanges []string) {
	// Read in the next however many values are in the queue and stop if it takes more than 10ms to do so
	nameColourQueue := make([]NameColourUpdate, 0)
	timeout := time.NewTimer(10_000_000)
	queueFlushed := false
	for !queueFlushed {
		select {
		case item := <-queue: // Still items remaining in the queue
			nameColourQueue = append(nameColourQueue, item)
		case <-timeout.C: // Either no more items remaining or took too long copying
			queueFlushed = true
		}
	}

	if len(nameColourQueue) == 0 {
		// Nothing to do
		return
	}

	// Build the batch from the items received earlier
	cellUpdateQueue := make([]*sheets.Request, 0)
	for _, item := range nameColourQueue {
		x, y, err := findNameCell(sheetsService, sheetID, item.Name, sheetRanges...)
		if err != nil {
			// Couldn't find the name, skip it
			continue
		}

		// Parse the colour string
		r, g, b, err := hexToRGB(item.Colour)
		if err != nil {
			// Failed to convert colour, complain and continue
			fmt.Fprintf(os.Stderr, "failed to convert hex colour to RGB: %v\n", err)
			continue
		}

		cellUpdateQueue = append(cellUpdateQueue, &sheets.Request{
			RepeatCell: &sheets.RepeatCellRequest{
				Cell: &sheets.CellData{
					UserEnteredFormat: &sheets.CellFormat{
						BackgroundColor: &sheets.Color{
							Red:   float64(r / 255),
							Green: float64(g / 255),
							Blue:  float64(b / 255),
						},
					},
				},
				Fields: "UserEnteredFormat(BackgroundColor)",
				Range: &sheets.GridRange{
					StartColumnIndex: int64(x),
					StartRowIndex:    int64(y),
					EndColumnIndex:   int64(x + 1),
					EndRowIndex:      int64(y + 1),
					SheetId:          0,
				},
			},
		})
	}

	if len(cellUpdateQueue) > 0 {
		_, err := sheetsService.Spreadsheets.BatchUpdate(sheetID, &sheets.BatchUpdateSpreadsheetRequest{
			IncludeSpreadsheetInResponse: false,
			Requests:                     cellUpdateQueue,
		}).Do()
		if err != nil {
			fmt.Fprintf(os.Stderr, "failed to submit batch update: %v\n", err)
		}
	}
}

// hexToRGB converts the given hexidecimal colour value string to its component red, green, and blue values.
func hexToRGB(hex string) (int, int, int, error) {
	values, err := strconv.ParseUint(hex, 16, 32)

	if err != nil {
		return 0, 0, 0, err
	}

	return int(values >> 16), int((values >> 8) & 0xFF), int(values & 0xFF), nil
}

// findNameCell attempts to find a name within the spreadsheet identified by the given sheetID and constrained to the given range(s).
// Returns the x and y coordinates of the matching cell.
func findNameCell(sheets *sheets.Service, sheetID string, name string, ranges ...string) (int, int, error) {
	result, err := sheets.Spreadsheets.Values.BatchGet(sheetID).Ranges(ranges...).Do()
	if err != nil {
		return 0, 0, fmt.Errorf("failed to get ranges from spreadsheet: %v", err)
	}

	for _, valueRange := range result.ValueRanges {
		// Get the x and y offset of this range
		xOffset, yOffset, _, _, err := parseA1Notation(valueRange.Range)
		if err != nil {
			return 0, 0, fmt.Errorf("failed parsing A1 range from result: %v", err)
		}

		for majorIndex, majorDimension := range valueRange.Values {
			for minorIndex, cell := range majorDimension {
				if cellValue, isString := cell.(string); isString && cellValue != "" && strings.Contains(name, cellValue) {
					// Found a match
					// Now figure out whether we were iterating horizontally or vertically to apply the appropriate
					// offsets
					if valueRange.MajorDimension == "COLUMNS" {
						// Outer iterator was per column
						return xOffset + majorIndex, yOffset + minorIndex, nil
					} else /*if valueRange.MajorDimension == "ROWS"*/ { // Always default to rows as this is standard
						// Outer iterator was per row
						return xOffset + minorIndex, yOffset + majorIndex, nil
					}
				}
			}
		}
	}

	return 0, 0, errors.New("unable to find matching cell")
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

// pluralise returns singular if n is 1, otherwise plural
func pluralise(singular string, plural string, n int) string {
	if n == 1 {
		return singular
	} else {
		return plural
	}
}

// discordInteractionRespond responds to the given interaction with the given message.
func discordInteractionRespond(session *discordgo.Session, interaction *discordgo.Interaction, message *string) {
	err := session.InteractionRespond(interaction, &discordgo.InteractionResponse{
		Type: discordgo.InteractionResponseChannelMessageWithSource,
		Data: &discordgo.InteractionResponseData{
			Content: *message,
		},
	})
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed responding to interaction: %v\n", err)
	}
}

// discordInteractionUpdate updates a response to the given interaction with the given message
func discordInteractionUpdate(session *discordgo.Session, interaction *discordgo.Interaction, message *string) {
	_, err := session.InteractionResponseEdit(interaction, &discordgo.WebhookEdit{
		Content: message,
	})
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed updating interaction response: %v\n", err)
	}
}

// Registers Discord event handlers
func registerDiscordEvents(session *discordgo.Session, config *Config, updateQueue chan<- NameColourUpdate) error {
	// State our intents for the events we want to receive
	session.Identify.Intents = discordgo.IntentGuildMessages | discordgo.IntentGuildMessageReactions

	// Create slash commands and register associated handlers
	for _, commandHandlerPair := range commands {
		_, err := session.ApplicationCommandCreate(config.Discord.ApplicationID, config.Discord.BotServer, commandHandlerPair.Command)
		if err != nil {
			return fmt.Errorf("failed registering command \"%v\": %v", "help", err)
		}

		session.AddHandler(commandHandlerPair.Handler)

		fmt.Printf("Registered \"%s\" command\n", commandHandlerPair.Command.Name)
	}

	// Register emote added
	session.AddHandler(func(s *discordgo.Session, e *discordgo.MessageReactionAdd) {
		// Ignore our own emoji events
		if session.State.User.ID == e.Member.User.ID {
			return
		}

		if value, exists := (*config).Discord.ReactionColours[e.Emoji.Name]; exists {
			fmt.Printf("%s reacted with '%s', changing their cell to %s\n", e.Member.Nick, e.MessageReaction.Emoji.Name, value.Colour)

			// Add name to the update queue
			updateQueue <- NameColourUpdate{
				Name:   e.Member.Nick,
				Colour: value.Colour,
			}
		} else {
			// No matching emoji, nothing to do
			fmt.Fprintf(os.Stderr, "%s reacted with unsupported emoji '%s'\n", e.Member.Nick, e.MessageReaction.Emoji.Name)
			return
		}
	})

	// Register emote removed
	session.AddHandler(func(s *discordgo.Session, e *discordgo.MessageReactionRemove) {
		member, err := s.GuildMember(e.GuildID, e.UserID)
		if err != nil {
			// Can't do anything without a user ID
			return
		}

		// Get the message to check if there are any other emoji from this user
		message, err := s.ChannelMessage(e.ChannelID, e.MessageID)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Failed getting message that emoji was removed from: %v", err)
			return
		}

		// Try pull out the next best matching colour from the reaction colour list, if there is one
		lowestPriority := -1
		matchedEmoji := ""
		matchedColour := "FFFFFF" // Default to white
		for _, reaction := range message.Reactions {
			// Reactions are returned in the order of occurrence on the message
			// Skip reactions that we don't have colours for
			var colourPriority ColourPriority
			if value, exists := config.Discord.ReactionColours[reaction.Emoji.Name]; exists {
				colourPriority = value
			} else {
				fmt.Fprintf(os.Stderr, "Unknown emoji '%s'\n", reaction.Emoji.Name)
				continue
			}

			// Get the users that have reacted
			users, err := s.MessageReactions(message.ChannelID, message.ID, reaction.Emoji.Name, 100, "", "")
			if err != nil {
				// Couldn't get the user list
				fmt.Fprintf(os.Stderr, "Failed getting list of users who reacted with a '%s' emoji: %v\n", reaction.Emoji.Name, err)
				continue
			}

			// Check if the user has reacted to this reaction as well
			// If it's a lower priority than the last one we'd found, then use this as the new colour
			for _, reactionUser := range users {
				if reactionUser.ID == member.User.ID {
					if colourPriority.Priority < lowestPriority || lowestPriority == -1 {
						lowestPriority = colourPriority.Priority
						matchedColour = colourPriority.Colour
						matchedEmoji = reaction.Emoji.Name
						break
					}
				}
			}
		}

		if matchedEmoji != "" {
			fmt.Printf("%s removed their '%s' react, but they still have a '%s' react. Changing their cell to %s\n", member.Nick, e.MessageReaction.Emoji.Name, matchedEmoji, matchedColour)
		} else {
			fmt.Printf("%s removed their '%s' react, changing their cell to %s\n", member.Nick, e.MessageReaction.Emoji.Name, matchedColour)
		}

		// Add name to the update queue
		updateQueue <- NameColourUpdate{
			Name:   member.Nick,
			Colour: matchedColour,
		}

	})

	// Register all emotes removed
	session.AddHandler(func(s *discordgo.Session, e *discordgo.MessageReactionRemoveAll) {
		user, err := s.User(e.UserID)
		if err != nil {
			// Can't do anything without a user ID
			return
		}

		// Add name to the update queue
		updateQueue <- NameColourUpdate{
			Name:   user.Username,
			Colour: "FFFFFF",
		}
	})

	return nil
}

// assertApplicationCommand returns whether the given interaction is an application command and the name of the command
// matches the one given.
func assertApplicationCommand(interaction *discordgo.InteractionCreate, name string) bool {
	return interaction.Type == discordgo.InteractionApplicationCommand && interaction.ApplicationCommandData().Name == name
}

// Initialises a sheets.Service using the given credentials and token paths.
func initSheets(credentialsPath string, tokenPath string) (*sheets.Service, error) {
	ctx := context.Background()
	b, err := os.ReadFile(credentialsPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read client secret file: %v", err)
	}

	oauthConfig, err := google.ConfigFromJSON(b, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to parse client secret file to config: %v", err)
	}
	client := getClient(tokenPath, oauthConfig)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve Sheets client: %v", err)
	}

	return srv, nil
}

// Initialises a discordgo.Session using the given bot token.
func initDiscord(token string) (*discordgo.Session, error) {
	discord, err := discordgo.New("Bot " + token)
	if err != nil {
		return discord, fmt.Errorf("unable to load Discord API: %v", err)
	}

	// Might as well open the session while we're at it
	err = discord.Open()
	if err != nil {
		return discord, fmt.Errorf("failed to open Discord session: %v", err)
	}

	return discord, nil
}
