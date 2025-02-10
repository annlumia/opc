//go:build windows
// +build windows

package browser

import (
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/konimarti/opc"
)

type Browser struct {
	nodes     []string
	programId string

	obj *ole.IDispatch
}

func NewBrowser(nodes []string, programId string) (*Browser, error) {
	object := opc.NewAutomationObject()
	defer object.Close()
	_, err := object.TryConnect(programId, nodes)
	if err != nil {
		return nil, err
	}

	return &Browser{
		nodes:     nodes,
		programId: programId,
	}, nil
}

func (b *Browser) Browse(query Query) *ItemsTag {
	itemsTag := &ItemsTag{}

	ao := opc.NewAutomationObject()
	defer ao.Close()
	_, err := ao.TryConnect(b.programId, b.nodes)
	if err != nil {
		errors := []string{"Connection failed", err.Error()}
		return &ItemsTag{
			Errors: &errors,
		}
	}

	b.obj = ao.Object()

	// create browser
	brw, err := oleutil.CallMethod(b.obj, "CreateBrowser")
	if err != nil {
		errors := []string{"Cannot call method CreateBrowser", err.Error()}
		return &ItemsTag{Errors: &errors}
	}

	browser := brw.ToIDispatch()

	itemsTag.Items = make([]*Tag, 0)

	// move to root
	oleutil.MustCallMethod(browser, "MoveToRoot")

	queryPath := strings.ReplaceAll(query.Path, "\\", "/")
	pathItems := strings.Split(queryPath, "/")

	key := strings.Join(pathItems, "/")

	for _, item := range pathItems {
		if item == "" {
			continue
		}
		oleutil.MustCallMethod(browser, "MoveDown", item)
	}

	var count int32

	// loop through branches
	oleutil.MustCallMethod(browser, "ShowBranches").ToIDispatch()
	count = oleutil.MustGetProperty(browser, "Count").Value().(int32)
	itemsTag.Total = uint16(count)

	if query.Offset > uint16(count) {
		errors := []string{"Offset is greater than total items"}
		itemsTag.Errors = &errors
		return itemsTag
	}

	max := int(query.Offset) + int(query.Limit)
	if max > int(count) {
		max = int(count)
	}

	for i := 1 + int(query.Offset); i <= max; i++ {
		itemName := oleutil.MustCallMethod(browser, "Item", i).Value()
		tag := &Tag{
			Name:     itemName.(string),
			IsBranch: true,
			Path:     key + "/" + itemName.(string),
		}
		itemsTag.Items = append(itemsTag.Items, tag)
	}

	// loop through leafs
	oleutil.MustCallMethod(browser, "ShowLeafs").ToIDispatch()
	count = oleutil.MustGetProperty(browser, "Count").Value().(int32)
	itemsTag.Total += uint16(count)
	max = int(query.Limit) - len(itemsTag.Items)
	if max > int(count) {
		max = int(count)
	}

	offset := int(query.Offset) - len(itemsTag.Items)
	if offset < 0 {
		offset = 0
	}

	for i := 1 + offset; i <= max; i++ {
		itemName := oleutil.MustCallMethod(browser, "Item", i).Value()
		itemId := oleutil.MustCallMethod(browser, "GetItemID", itemName).Value()

		tag := &Tag{
			Name:     itemName.(string),
			ItemId:   itemId.(string),
			IsBranch: false,
			Path:     key + "/" + itemName.(string),
		}
		itemsTag.Items = append(itemsTag.Items, tag)

		if len(itemsTag.Items) == int(query.Limit) {
			break
		}
	}

	return itemsTag
}
