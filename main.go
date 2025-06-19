package main

import (
	"fmt"
	"log"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type GradeEntry struct {
	Date        string
	Subject     string
	Topic       string
	Weight      float64
	Result      float64
	VerbalGrade string
}

type SubjectGradeData struct {
	TotalWeightedSum float64
	TotalWeight      float64
	GradeCount       int
	Average          float64
}

func ProcessGradesFromFile(filePath string) (map[string]*SubjectGradeData, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("error opening Excel file %s: %w", filePath, err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("Error closing Excel file: %v", err)
		}
	}()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, fmt.Errorf("could not find any sheets in the XLSX file")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("error getting rows from sheet '%s': %w", sheetName, err)
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("no grade data found in the XLSX file (expected at least 2 rows, including headers)")
	}

	header := rows[0]
	subjectColIdx := -1
	weightColIdx := -1
	resultColIdx := -1

	for idx, cell := range header {
		trimmedCell := strings.TrimSpace(cell)
		switch trimmedCell {
		case "Předmět":
			subjectColIdx = idx
		case "Váha":
			weightColIdx = idx
		case "Výsledek":
			resultColIdx = idx
		}
	}

	if subjectColIdx == -1 || weightColIdx == -1 || resultColIdx == -1 {
		return nil, fmt.Errorf("no valid grade data found to process from the Excel file. Make sure your 'Předmět', 'Váha' and 'Výsledek' columns contain valid numeric entries")
	}

	subjectGrades := make(map[string]*SubjectGradeData)
	processedGradesCount := 0

	for rowIdx, row := range rows {
		if rowIdx == 0 {
			continue
		}

		if len(row) <= subjectColIdx || len(row) <= weightColIdx || len(row) <= resultColIdx {
			log.Printf("Warning: Skipping row %d due to insufficient columns for required data (found %d, expected at least %d).", rowIdx+1, len(row), max(subjectColIdx, weightColIdx, resultColIdx)+1)
			continue
		}

		subject := strings.TrimSpace(row[subjectColIdx])
		weightStr := row[weightColIdx]
		resultStr := row[resultColIdx]

		if subject == "" {
			log.Printf("Warning: Skipping row %d due to empty subject name.", rowIdx+1)
			continue
		}

		weight, errW := strconv.ParseFloat(weightStr, 64)
		result, errR := strconv.ParseFloat(resultStr, 64)

		if errW != nil || errR != nil {
			log.Printf("Warning: Skipping row %d for subject '%s' due to non-numeric grade/weight: Weight='%s' (err: %v), Result='%s' (err: %v)", rowIdx+1, subject, weightStr, errW, resultStr, errR)
			continue
		}

		data, exists := subjectGrades[subject]
		if !exists {
			data = &SubjectGradeData{}
			subjectGrades[subject] = data
		}

		data.TotalWeightedSum += result * weight
		data.TotalWeight += weight
		data.GradeCount++
		processedGradesCount++
	}

	if processedGradesCount == 0 {
		return nil, fmt.Errorf("no valid grade data found to process from the Excel file. Make sure your 'Předmět', 'Váha' and 'Výsledek' columns contain valid numeric entries")
	}

	for _, data := range subjectGrades {
		if data.GradeCount > 0 && data.TotalWeight > 0 {
			data.Average = data.TotalWeightedSum / data.TotalWeight
		} else {
			log.Printf("Warning: Subject data has zero grades or zero total weight, average cannot be calculated.")
			data.Average = 0
		}
	}

	return subjectGrades, nil
}

func main() {
	inputFilePath := "PodepsaniHodnoceni.xlsx"

	subjectAverages, err := ProcessGradesFromFile(inputFilePath)
	if err != nil {
		log.Fatalf("Error processing grades: %v", err)
	}

	if len(subjectAverages) == 0 {
		fmt.Println("No subject averages could be calculated from the provided data.")
		return
	}

	fmt.Println("--- Subject Weighted Averages ---")

	var subjectsSorted []string
	for subject := range subjectAverages {
		subjectsSorted = append(subjectsSorted, subject)
	}
	sort.Strings(subjectsSorted)

	for _, subject := range subjectsSorted {
		data := subjectAverages[subject]
		fmt.Printf("%s: %.3f\n", subject, data.Average)
	}
}
