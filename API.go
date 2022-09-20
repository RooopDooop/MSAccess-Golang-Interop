package main

import (
	"database/sql"
	"encoding/json"
	"fmt"
	_ "go/format"
	"log"
	"net/http"

	_ "github.com/denisenkom/go-mssqldb"
	"github.com/gorilla/mux"
)

type Order struct {
	OrderID             string  `json:"OrderID"`
	ProductionLotNumber *string `json:"productionLotNumber"`
	StopNumber          int     `json:"StopNumber"`
}

type OrderSearch struct {
	OrderID               string  `json:"orderID"`
	TotalList             *string `json:"totalList"`
	StaffCode             *string `json:"staffCode"`
	ClearedByMRP          *string `json:"ClearedByMRP"`
	Tag                   *string `json:"tag"`
	OnHoldReason          *string `json:"onHoldReason"`
	CSR                   *string `json:"csr"`
	Complexity            *string `json:"complexity"`
	PDR                   *string `json:"pdr"`
	Project               *string `json:"project"`
	ProductionLotNumber   *string `json:"productionLotNumber"`
	OrderDate             *string `json:"orderDate"`
	TotalNet              *string `json:"totalNet"`
	FaxOrderNotes         *string `json:"faxOrderNotes"`
	LastOfTransactionNote *string `json:"lastOfTransactionNote"`
	RunOnMRP              *string `json:"runOnMRP"`
}

var db *sql.DB
var err error

var ProdOrderList []Order

func main() {
	fmt.Println("Connecting to Sandbox DB...")
	db, err = sql.Open("mssql", "odbc:server=SERVER;database=DATABASE;")
	fmt.Println("Connected.")

	handleRequests()
}

func handleRequests() {
	myRouter := mux.NewRouter().StrictSlash(true)

	myRouter.HandleFunc("/GetProductionOrders", returnProductionOrders)
	myRouter.HandleFunc("/GetOrderSearches", returnOrderQuotesOnload)

	log.Fatal(http.ListenAndServe(":10000", myRouter))
}

func returnProductionOrders(w http.ResponseWriter, request *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type,access-control-allow-origin, access-control-allow-headers")
	w.Header().Set("Content-Type", "application/json")

	fmt.Println("Endpoint Hit: returnSpecificOrder")

	sql := "SELECT dbo.Orders.OrderID, " +
		"dbo.Orders.productionLotNumber, " +
		"dbo.Orders.StopNumber " +
		"FROM dbo.Orders " +
		"INNER JOIN dbo.companyStaff ON dbo.Orders.StaffCode = dbo.companyStaff.StaffCode " +
		"INNER JOIN dbo.companyBranch ON dbo.companyStaff.BranchCode = dbo.companyBranch.BranchCode " +
		"INNER JOIN dbo.Employees ON dbo.Orders.EmployeeID = dbo.Employees.EmployeeID " +
		"INNER JOIN dbo.lotSchedule ON dbo.Orders.productionLotNumber = dbo.lotSchedule.lot " +
		"GROUP BY dbo.Orders.OrderID, " +
		"dbo.Orders.productionLotNumber, " +
		"dbo.Orders.StopNumber, " +
		"dbo.companyBranch.BranchName, " +
		"dbo.Employees.FirstName, " +
		"dbo.Employees.LastName, " +
		"dbo.Orders.Tag, " +
		"dbo.lotSchedule.lot, " +
		"dbo.Orders.OnHoldReason, " +
		"dbo.Orders.CabinetStatus, " +
		"dbo.Orders.DeskStatus, " +
		"dbo.Orders.HutchStatus, " +
		"dbo.Orders.TableStatus, " +
		"dbo.Orders.WallunitStatus " +
		"HAVING ((dbo.Orders.OnHoldReason = 'In Production') OR " +
		"(dbo.Orders.OnHoldReason = 'Hold,Hardware') OR " +
		"(dbo.Orders.OnHoldReason = 'Hold,Material') OR " +
		"(dbo.Orders.OnHoldReason = 'Hold,Delivery')) " +
		"ORDER BY dbo.Orders.productionLotNumber, dbo.Orders.StopNumber ASC;"

	rows, err := db.Query(sql)

	if err == nil {
		listOrders := []Order{}

		for rows.Next() {
			var r Order

			rows.Scan(&r.OrderID,
				&r.ProductionLotNumber,
				&r.StopNumber)

			listOrders = append(listOrders, r)
		}
		listColumns := []string{"OrderID", "productionLotNumber", "StopNumber"}
		json.NewEncoder(w).Encode(map[string]interface{}{"Columns": listColumns, "Results": listOrders})
	} else {
		w.WriteHeader(http.StatusInternalServerError)
		fmt.Println("Error detected, problem in returnSpecificOrder")
	}
}

func returnOrderQuotesOnload(w http.ResponseWriter, request *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type,access-control-allow-origin, access-control-allow-headers")
	w.Header().Set("Content-Type", "application/json")

	fmt.Println("Endpoint Hit: returnSpecificOrder")

	sql := "SELECT DISTINCT TOP(50000) dbo.Orders.OrderID, " +
		"dbo.Orders.TotalList, " +
		"dbo.Orders.StaffCode, " +
		"dbo.Orders.clearedByMRP, " +
		"dbo.Orders.Tag, " +
		"dbo.Orders.OnHoldReason, " +
		"dbo.Orders.employeeID AS CSR, " +
		"dbo.Orders.Complexity, " +
		"dbo.Orders.pdr, " +
		"dbo.Orders.RegistrationNumber AS Project, " +
		"dbo.Orders.productionLotNumber, " +
		"dbo.Orders.OrderDate, " +
		"dbo.Orders.TotalList*(1-dbo.Orders.Discount) AS TotalNet, " +
		"dbo.Orders.faxOrderNotes, " +
		"(SELECT TOP(1) transactionNote FROM dbo.bisTransactionLog as SubTrans WHERE SubTrans.transactionType = 99 AND SubTrans.transactionKey = dbo.bisTransactionLog.transactionKey Order by transactionNote DESC) as LastOfTransactionNote, " +
		"dbo.bisTransactionLog.transactionKey AS runOnMRP " +
		"FROM dbo.Orders " +
		"LEFT JOIN dbo.bisTransactionLog ON dbo.Orders.OrderID = dbo.bisTransactionLog.transactionKey " +
		"LEFT JOIN dbo.LotHardware ON dbo.Orders.OrderID = dbo.LotHardware.orderID " +
		"WHERE dbo.Orders.OrderDate > '2019' " +
		"Order By OrderDate desc"

	rows, err := db.Query(sql)

	if err == nil {
		listOrderSearch := []OrderSearch{}

		for rows.Next() {
			var o OrderSearch

			rows.Scan(&o.OrderID,
				&o.TotalList,
				&o.StaffCode,
				&o.ClearedByMRP,
				&o.Tag,
				&o.OnHoldReason,
				&o.CSR,
				&o.Complexity,
				&o.PDR,
				&o.Project,
				&o.ProductionLotNumber,
				&o.OrderDate,
				&o.TotalNet,
				&o.FaxOrderNotes,
				&o.LastOfTransactionNote,
				&o.RunOnMRP)

			listOrderSearch = append(listOrderSearch, o)
		}
		json.NewEncoder(w).Encode(map[string]interface{}{"Results": listOrderSearch})
	} else {
		w.WriteHeader(http.StatusInternalServerError)
		fmt.Println("Error detected, problem in returnSpecificOrder")
	}
}
