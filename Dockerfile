FROM alpine:latest AS build

RUN apk add --no-cache go

WORKDIR /build/
COPY go.mod go.sum *.go ./

RUN go build -o rolly

# Copy over the built binary
FROM alpine:latest

WORKDIR /app/
COPY --from=build /build/rolly ./rolly

ENTRYPOINT ["/app/rolly"]
