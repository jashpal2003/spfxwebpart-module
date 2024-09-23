import * as React from 'react';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { IEmployeeListingProps } from './IEmployeeListingProps';
import { IColumn, DetailsList, DetailsListLayoutMode, TextField, PrimaryButton, IconButton, Dialog, DialogType, DialogFooter, DatePicker, Dropdown, IDropdownOption } from '@fluentui/react';

sp.setup({
  sp: {
    baseUrl: "https://t12y7.sharepoint.com/sites/MySite101"
  }
});

interface IEmployee {
  FirstName: string;
  DepartmentId: number;
  Experience: string;
  DOB: string;
  Id: number;
}

interface IEmployeeListingState {
  employees: IEmployee[];
  columns: IColumn[];
  searchQuery: string;
  filteredEmployees: IEmployee[];
  showEditDialog: boolean;
  showAddDialog: boolean;
  selectedEmployee: IEmployee | undefined;
  departments: IDropdownOption[];
  showConfirmationDialog: boolean;
  departmentMap: { [key: number]: string }; // Add departmentMap
}

export default class EmployeeListing extends React.Component<IEmployeeListingProps, IEmployeeListingState> {
  constructor(props: IEmployeeListingProps) {
    super(props);
    this.state = {
      employees: [],
      columns: this._getColumns(),
      searchQuery: '',
      filteredEmployees: [],
      showEditDialog: false,
      showAddDialog: false,
      selectedEmployee: undefined,
      departments: [],
      showConfirmationDialog: false,
      departmentMap: {} // Initialize departmentMap
    };
  }

  public componentDidMount(): void {
    this._fetchEmployees();
    this._fetchDepartments();
  }

  private _fetchEmployees(): void {
    sp.web.lists.getByTitle("Employees").items.select("Id", "FirstName", "DepartmentId", "Experience", "DOB").get().then((items: IEmployee[]) => {
      this.setState({ employees: items, filteredEmployees: items });
    }).catch((error: Error) => {
      console.error("Error fetching employees:", error);
    });
  }

  private _fetchDepartments(): void {
    sp.web.lists.getByTitle("Department List").items.select("Id", "Title").get().then((items) => {
      const departments = items.map((item) => ({ key: item.Id.toString(), text: item.Title }));
      const departmentMap = items.reduce((map, item) => {
        map[item.Id] = item.Title;
        return map;
      }, {} as { [key: number]: string });
      this.setState({ departments, departmentMap });
    }).catch((error: Error) => {
      console.error("Error fetching departments:", error);
    });
  }

  private _getColumns(): IColumn[] {
    return [
      {
        key: 'Actions',
        name: 'Actions',
        fieldName: '',
        minWidth: 100,
        maxWidth: 100,
        onRender: (item: IEmployee) => (
          <div>
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete"
              onClick={() => this._deleteEmployee(item.Id)}
              styles={{ root: { marginRight: 10 } }}
            />
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit"
              onClick={() => this._openEditDialog(item)}
            />
          </div>
        )
      },
      {
        key: 'FirstName',
        name: 'FirstName',
        fieldName: 'FirstName',
        minWidth: 100,
        maxWidth: 200,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick
      },
      {
        key: 'Department',
        name: 'Department',
        fieldName: 'DepartmentId',
        minWidth: 100,
        maxWidth: 200,
        onRender: (item: IEmployee) => {
          const departmentName = this.state.departmentMap[item.DepartmentId];
          console.log("-=-=-=-=-=-=-=",departmentName);
          return departmentName || '';
        }
      },
      {
        key: 'Experience',
        name: 'Experience',
        fieldName: 'Experience',
        minWidth: 100,
        maxWidth: 200
      },
      {
        key: 'DOB',
        name: 'DOB',
        fieldName: 'DOB',
        minWidth: 100,
        maxWidth: 200
      }
    ];
  }

  private _deleteEmployee = (id: number): void => {
    sp.web.lists.getByTitle("Employees").items.getById(id).delete().then(() => {
      this._fetchEmployees();
    }).catch((error: Error) => {
      console.error("Error deleting employee:", error);
    });
  };

  private _openEditDialog = (employee: IEmployee): void => {
    this.setState({ showEditDialog: true, selectedEmployee: employee });
  };

  private _openAddDialog = (): void => {
    this.setState({ showAddDialog: true, selectedEmployee: { FirstName: '', DepartmentId: 0, Experience: '', DOB: '', Id: 0 } });
  };

  private _closeEditDialog = (): void => {
    this.setState({ showEditDialog: false, selectedEmployee: undefined });
  };

  private _closeAddDialog = (): void => {
    this.setState({ showAddDialog: false, selectedEmployee: undefined });
  };

  private _saveEditedEmployee = (): void => {
    this.setState({ showConfirmationDialog: true });
  };

  private _confirmSave = (): void => {
    const { selectedEmployee } = this.state;
    if (selectedEmployee) {
      if (selectedEmployee.Id === 0) {
        sp.web.lists.getByTitle("Employees").items.add({
          FirstName: selectedEmployee.FirstName,
          DepartmentId: selectedEmployee.DepartmentId,
          Experience: selectedEmployee.Experience,
          DOB: selectedEmployee.DOB
        }).then(() => {
          this._fetchEmployees();
          this._closeAddDialog();
        }).catch((error: Error) => {
          console.error("Error adding employee:", error);
        });
      } else {
        sp.web.lists.getByTitle("Employees").items.getById(selectedEmployee.Id).update({
          FirstName: selectedEmployee.FirstName,
          DepartmentId: selectedEmployee.DepartmentId,
          Experience: selectedEmployee.Experience,
          DOB: selectedEmployee.DOB
        }).then(() => {
          this._fetchEmployees();
          this._closeEditDialog();
        }).catch((error: Error) => {
          console.error("Error updating employee:", error);
        });
      }
    }
    this.setState({ showConfirmationDialog: false });
  };

  private _cancelSave = (): void => {
    this.setState({ showConfirmationDialog: false });
  };

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { filteredEmployees } = this.state;
    const newColumns: IColumn[] = this._getColumns();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];

    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = false;
      }
    });

    const sortedEmployees = this._copyAndSort(filteredEmployees, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      filteredEmployees: sortedEmployees
    });
  };

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  private _onSearch = (): void => {
    const { employees, searchQuery } = this.state;
    const filteredEmployees = employees.filter(employee =>
      employee.FirstName.toLowerCase().includes(searchQuery.toLowerCase())
    );
    this.setState({ filteredEmployees });
  };

  private _onSearchQueryChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ searchQuery: newValue || '' });
  };

  private _onEmployeeFieldChange = (fieldName: keyof IEmployee, value: string | number | Date | { Id: number }): void => {
    const { selectedEmployee } = this.state;
    if (selectedEmployee) {
      this.setState({
        selectedEmployee: {
          ...selectedEmployee,
          [fieldName]: value
        }
      });
    }
  };

  public render(): React.ReactElement<{}> {
    const { columns, filteredEmployees, searchQuery, showEditDialog, showAddDialog, selectedEmployee, departments, showConfirmationDialog } = this.state;
    return (
      <div>
        <TextField
          label="Search by Name"
          value={searchQuery}
          onChange={this._onSearchQueryChange}
          styles={{ root: { marginBottom: 10 } }}
        />
        <PrimaryButton text="Search" onClick={this._onSearch} styles={{ root: { marginBottom: 10 } }} />
        <PrimaryButton text="Add Employee" onClick={this._openAddDialog} styles={{ root: { marginBottom: 10 } }} />
        <DetailsList
          items={filteredEmployees}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick
        />
        {showEditDialog && selectedEmployee && (
          <Dialog
            hidden={!showEditDialog}
            onDismiss={this._closeEditDialog}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Edit Employee'
            }}
            modalProps={{
              isBlocking: true
            }}
          >
            <TextField
              label="First Name"
              value={selectedEmployee.FirstName}
              onChange={(e, newValue) => this._onEmployeeFieldChange('FirstName', newValue || '')}
            />
            <Dropdown
              label="Department"
              selectedKey={selectedEmployee.DepartmentId.toString()}
              options={departments}
              onChange={(e, option) => this._onEmployeeFieldChange('DepartmentId', parseInt(option?.key as string))}
            />
            <TextField
              label="Experience"
              value={selectedEmployee.Experience}
              onChange={(e, newValue) => this._onEmployeeFieldChange('Experience', newValue || '')}
            />
            <DatePicker
              label="DOB"
              value={selectedEmployee.DOB ? new Date(selectedEmployee.DOB) : undefined}
              onSelectDate={(date) => {
                console.log('Selected date:', date);
                this._onEmployeeFieldChange('DOB', date?.toISOString() || '');
              }}
            />
            <DialogFooter>
              <PrimaryButton onClick={this._saveEditedEmployee} text="Save" />
              <PrimaryButton onClick={this._closeEditDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        )}
        {showAddDialog && selectedEmployee && (
          <Dialog
            hidden={!showAddDialog}
            onDismiss={this._closeAddDialog}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Add Employee'
            }}
            modalProps={{
              isBlocking: true
            }}
          >
            <TextField
              label="First Name"
              value={selectedEmployee.FirstName}
              onChange={(e, newValue) => this._onEmployeeFieldChange('FirstName', newValue || '')}
            />
            <Dropdown
              label="Department"
              selectedKey={selectedEmployee.DepartmentId.toString()}
              options={departments}
              onChange={(e, option) => this._onEmployeeFieldChange('DepartmentId', parseInt(option?.key as string))}
            />
            <TextField
              label="Experience"
              value={selectedEmployee.Experience}
              onChange={(e, newValue) => this._onEmployeeFieldChange('Experience', newValue || '')}
            />
            <DatePicker
              label="DOB"
              value={selectedEmployee.DOB ? new Date(selectedEmployee.DOB) : undefined}
              onSelectDate={(date) => {
                console.log('Selected date:', date);
                this._onEmployeeFieldChange('DOB', date?.toISOString() || '');
              }}
            />
            <DialogFooter>
              <PrimaryButton onClick={this._saveEditedEmployee} text="Save" />
              <PrimaryButton onClick={this._closeAddDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        )}
        {showConfirmationDialog && (
          <Dialog
            hidden={!showConfirmationDialog}
            onDismiss={this._cancelSave}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Confirm Save',
              closeButtonAriaLabel: 'Close',
              subText: 'Are you sure you want to save these changes?'
            }}
            modalProps={{
              isBlocking: true
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={this._confirmSave} text="Yes" />
              <PrimaryButton onClick={this._cancelSave} text="No" />
            </DialogFooter>
          </Dialog>
        )}
      </div>
    );
  }
}
