package example

type User struct {
	Name   string
	Age    int
	Height int
}

type UserWithTag struct {
	Name   string `xlsx:"1"`
	Age    int    `xlsx:"2"`
	Height int    `xlsx:"3"`
}

type UserWithTag2 struct {
	Name   string `xlsx:"1"`
	Age    int    `xlsx:"3"`
	Height int    `xlsx:"5"`
}
