package browser

type ItemsTag struct {
	Items []*Tag `json:"items,omitempty"`

	Total  uint16 `json:"total,omitempty"`
	Offset uint16 `json:"offset,omitempty"`
	Limit  uint16 `json:"limit,omitempty"`

	Errors *[]string `json:",omitempty"`
}
