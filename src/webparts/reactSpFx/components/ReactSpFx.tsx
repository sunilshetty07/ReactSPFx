import * as React from 'react';
//import styles from './ReactSpFx.module.scss';
import type { IReactSpFxProps } from './IReactSpFxProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP } from '../ReactSpFxWebPart';
import { DefaultButton, DetailsList, Dropdown, IColumn, IDropdownOption, INavLink, INavLinkGroup, INavStyles, Nav, PrimaryButton, TextField } from '@fluentui/react';
import { Col, Row } from 'reactstrap';
import 'bootstrap/dist/css/bootstrap.min.css';
import { useEffect, useState } from 'react';
import SearchItems from './SearchItems';
//import { ListInlineItem } from 'reactstrap';

interface IListItem {
  Id: number;
  Title: string;
  Created: string;
  Author?: string;
}

const ReactSpFx: React.FC<IReactSpFxProps> = (props: IReactSpFxProps) => {
  const [items, setItems] = useState<IListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [title, setTitle] = React.useState("");
  const [navItems, setNavItems] = React.useState("");
  
  const sp: SPFI = getSP(props.context);
  let listName = props.selectedList;


  useEffect(() => {
    const fetchItems = async () => {
      console.log("Selected List in useEffect:", listName);
      try {
        const listItems: any[] = await sp.web.lists
          .getByTitle(listName)
          .items.select("Id", "Title", "Created", "Author/Title")
          .expand("Author")();
        const formattedItems: IListItem[] = listItems.map(item => ({
          Id: item.Id,
          Title: item.Title,
          Created: new Date(item.Created).toLocaleDateString(),
          Author: item.Author?.Title
        }));

        setItems(formattedItems);
      } catch (error) {
        console.error("Error fetching items:", error);
      } finally {
        setLoading(false);
      }
    };

    fetchItems();
  }, [listName]);

  const dropdownOptions: IDropdownOption[] = items.map(item => ({
    key: item.Id,
    text: `${item.Id} - ${item.Title}`,
  }));
  const onSelectItem = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const selectedItem = items.find(i => i.Id === option.key);
      setSelectedId(option.key as number);
      setTitle(selectedItem?.Title || "");
    }
  };

  function _onLinkClick(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) {
    ev?.preventDefault();
    if (item) {
      setNavItems(item.name);
      console.log(navItems); // update state
      //alert(`${item.name} link clicked`); // use clicked value directly
    }
  }
  const navStyles: Partial<INavStyles> = {
    root: {
      width: '100%',
      height: '100%',
      boxSizing: 'border-box',
      border: '1px solid #eee',
      overflowY: 'auto',
    },
  };
  const navLinkGroups: INavLinkGroup[] = [
    {
      links: [
        {
          name: 'AllItems',
          url: '#',
          key: 'key3'
        },
        {
          name: 'New Item',
          url: '#',
          key: 'key4',
        },
        {
          name: 'Update Item',
          url: '#',
          key: 'key5'
        },
        {
          name: 'Delete Item',
          url: '#',
          key: 'key6'
        },
        {
          name: 'Search Item',
          url: '#',
          key: 'key7'
        }
      ],
    },
  ];
  const columns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 50, maxWidth: 70, isResizable: true },
    { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'created', name: 'Created', fieldName: 'Created', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'author', name: 'Author', fieldName: 'Author', minWidth: 100, maxWidth: 150, isResizable: true },
  ];

  //add item
  const addItem = async () => {
    try {
      const result = await sp.web.lists.getByTitle(listName).items.add({
        Title: title,  // assuming your list has a Title column
      });
      console.log("Item added:", result);
      alert(`Item added successfully!${result.ID}`);
      setTitle(""); // clear input
    } catch (err) {
      console.error("Error adding item:", err);
      alert("Failed to add item.");
    }
  };

  // Update SharePoint list item
  const updateItem = async () => {
  if (selectedId !== null) {
    try {
      await sp.web.lists
        .getByTitle(listName)
        .items.getById(selectedId)
        .update({
          Title: title
        });

      alert("Item updated successfully!");

      // 🔄 Update the item in local state instead of removing it
      const updatedItems = items.map(i =>
        i.Id === selectedId ? { ...i, Title: title } : i
      );

      setItems(updatedItems);

      // 🔄 Reset selection and input
      setSelectedId(null);
      setTitle("");
    } catch (error) {
      console.error("Error updating item:", error);
      alert("Error updating item.");
    }
  }
};

  // Delete item
  const deleteItem = async () => {
    if (selectedId !== null) {
      const confirmDelete = confirm("Are you sure you want to delete this item?");
      if (!confirmDelete) return;

      try {
        await sp.web.lists.getByTitle(listName).items.getById(selectedId).delete();
        alert("Item deleted successfully!");

        // Refresh items after deletion
        const updatedItems = items.filter(i => i.Id !== selectedId);
        setItems(updatedItems);
        setSelectedId(null);
        setTitle("");
      } catch (error) {
        console.error("Error deleting item:", error);
        alert("Error deleting item.");
      }
    }
  };

  return (
    <>
      <Row>
        <Col
          className="bg-light border"
          xs="3"
        >
          <Nav
            onLinkClick={_onLinkClick}
            selectedKey={navItems}
            ariaLabel="Nav basic example"
            styles={navStyles}
            groups={navLinkGroups}
          />
        </Col>
        <Col className="bg-light border" xs="9">
          {navItems === "AllItems" ? (
            // If Items is selected
            <div style={{ padding: 20 }}>
              {loading ? (
                <div>Loading list items...</div>
              ) : (
                <DetailsList
                  items={items}
                  columns={columns}
                  setKey="set"
                  layoutMode={0}
                  isHeaderVisible={true}
                  selectionPreservedOnEmptyClick={true}
                  styles={{ root: { overflowX: 'auto' } }}
                />
              )}
            </div>
          ) : navItems === "New Item" ? (
            // If Pages is selected
            <div style={{ padding: 20 }}>
              <h2>📄 Create New Items</h2>
              <p>select add item button to create new item</p>
              <TextField
                label="Title"
                value={title}
                onChange={(_, newValue) => setTitle(newValue || "")}
              />
              <PrimaryButton text="Add Item" onClick={addItem} />
            </div>
          ) : navItems === "Update Item" ? (
            // If Notebook is selected
            <div style={{ padding: 20 }}>
              <Dropdown
                placeholder="Select an item"
                label="Select Item by ID"
                options={dropdownOptions}
                onChange={onSelectItem}
                selectedKey={selectedId}
                styles={{ dropdown: { width: 300 } }}
              />

              {selectedId && (
                <div style={{ marginTop: 20 }}>
                  <TextField
                    label="Title"
                    value={title}
                    onChange={(_, newValue) => setTitle(newValue || "")}
                    styles={{ root: { width: 300 } }}
                  />
                  <PrimaryButton
                    text="Update"
                    style={{ marginTop: 10 }}
                    onClick={updateItem}
                  />
                </div>
              )}
            </div>
          ) : navItems === "Delete Item" ? (
            // If Notebook is selected
            <div style={{ padding: 20 }}>
              <Dropdown
                placeholder="Select an item"
                label="Select Item by ID"
                options={dropdownOptions}
                onChange={onSelectItem}
                selectedKey={selectedId}
                styles={{ dropdown: { width: '75%' } }}
              />

              {selectedId && (
                <div style={{ marginTop: 20 }}>
                  <TextField
                    label="Title"
                    value={title}
                    disabled={true}
                    onChange={(_, newValue) => setTitle(newValue || "")}
                    styles={{ root: { width: '75%' } }}
                  />

                  <div style={{ marginTop: 10, display: "flex", gap: "10px" }}>
                    <DefaultButton text="Delete" onClick={deleteItem} />
                  </div>
                </div>)}
            </div>
          ) : navItems === "Search Item" ? (
            // If Notebook is selected
            <SearchItems listName={listName} context={props.context} />
          ):
            (
            // Optional: fallback for any other nav item
            <div>
              <span>Select navigation</span>
            </div>
          )}
        </Col>
      </Row>
    </>
  );

}
export default ReactSpFx;
