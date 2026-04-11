package config

import "os"

// Config holds server configuration loaded from environment variables.
type Config struct {
	ServerName      string
	Version         string
	LogLevel        string
	PythonPath      string
	EngineRoot      string
	MaxTasks        int
	MaxTaskLen      int
	AllowedHomeOnly bool
}

// Load reads configuration from environment variables with sensible defaults.
func Load() *Config {
	return &Config{
		ServerName:      getEnv("SERVER_NAME", "excel-engine-go"),
		Version:         getEnv("VERSION", "v1.1.0"),
		LogLevel:        getEnv("LOG_LEVEL", "info"),
		PythonPath:      getEnv("PYTHON_PATH", "python3"),
		EngineRoot:      getEnv("ENGINE_ROOT", ""),
		MaxTasks:        500,
		MaxTaskLen:      10000,
		AllowedHomeOnly: true,
	}
}

func getEnv(key, defaultValue string) string {
	if value := os.Getenv(key); value != "" {
		return value
	}
	return defaultValue
}
