# Stage 1: Build the Go application
FROM golang:1.20-alpine AS builder

# Set the working directory inside the container
WORKDIR /app

# Copy go.mod and go.sum files first (to cache dependencies)
COPY go.mod go.sum ./

# Download dependencies
RUN go mod download

# Copy the rest of the source code
COPY . .

# Build the Go application
RUN CGO_ENABLED=0 GOOS=linux go build -o GoExcel .

# Stage 2: Create a small runtime image
FROM alpine:latest

# Install CA certificates (if needed, for making HTTPS requests)
RUN apk --no-cache add ca-certificates

# Set the working directory
WORKDIR /root/

# Copy the built Go binary from the builder stage
COPY --from=builder /app/GoExcel .

# Expose the port that the service listens on
EXPOSE 10000

# Run the Go binary
CMD ["./GoExcel"]
