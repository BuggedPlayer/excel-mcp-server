package tools

import (
	"context"
	"fmt"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// WithRecovery wraps a tool handler function with panic recovery.
// If the handler panics, it returns an error result instead of crashing the server.
func WithRecovery(handler server.ToolHandlerFunc) server.ToolHandlerFunc {
	return func(ctx context.Context, request mcp.CallToolRequest) (result *mcp.CallToolResult, err error) {
		defer func() {
			if r := recover(); r != nil {
				err = fmt.Errorf("internal error: %v", r)
				result = nil
			}
		}()
		return handler(ctx, request)
	}
}
