package browser

type Query struct {
	Path   string `json:"path"`
	Offset uint16 `json:"offset"`
	Limit  uint16 `json:"limit"`
}
