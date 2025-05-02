using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using Inventory;
using Inventory.Model;

namespace Inventory.Server.Controllers
{
    public class ServersController : ApiController
    {
        private InventoryContext db = new InventoryContext();

        // GET: api/Servers
        public IQueryable<Model.Server> GetServers()
        {
            return db.Servers;
        }

        // GET: api/Servers/5
        [ResponseType(typeof(Model.Server))]
        public IHttpActionResult GetServer(string id)
        {
            Model.Server server = db.Servers.Find(id);
            if (server == null)
            {
                return NotFound();
            }

            return Ok(server);
        }

        // PUT: api/Servers/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutServer(string id, Model.Server server)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != server.Id)
            {
                return BadRequest();
            }

            db.Entry(server).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ServerExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Servers
        [ResponseType(typeof(Model.Server))]
        public IHttpActionResult PostServer(Model.Server server)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Servers.Add(server);

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateException)
            {
                if (ServerExists(server.Id))
                {
                    return Conflict();
                }
                else
                {
                    throw;
                }
            }

            return CreatedAtRoute("DefaultApi", new { id = server.Id }, server);
        }

        // DELETE: api/Servers/5
        [ResponseType(typeof(Model.Server))]
        public IHttpActionResult DeleteServer(string id)
        {
            Model.Server server = db.Servers.Find(id);
            if (server == null)
            {
                return NotFound();
            }
    
            db.Servers.Remove(server);
            db.SaveChanges();

            return Ok(server);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ServerExists(string id)
        {
            return db.Servers.Count(e => e.Id == id) > 0;
        }
    }
}