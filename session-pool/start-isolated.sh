#!/bin/bash
# Start M365 Session Pool in isolated mode (one container per connection)

cd "$(dirname "$0")"

echo "Starting M365 Session Pool (isolated mode)..."
echo "  - Router on port 5200"
echo "  - Each connection gets its own container"

docker compose -f docker-compose.isolated.yml up -d --build

echo ""
echo "Waiting for containers to start..."
sleep 5

echo ""
echo "Container status:"
docker compose -f docker-compose.isolated.yml ps

echo ""
echo "Health check:"
curl -s http://localhost:5200/health | jq .

echo ""
echo "To view status: curl http://localhost:5200/status"
echo "To view metrics: curl http://localhost:5200/metrics"
echo "To restart a connection: curl -X POST http://localhost:5200/container/ForIT-GA/restart"
